// Running one interpreter process: the part both platforms' runners share.
//
// Each platform decides how its script reaches the interpreter and how a
// failure reads in that interpreter's stderr. Everything else — the settings a
// run takes, telling an abort from a timeout from an overflow, and reporting
// the run to the debug hook — happens here, once.
import { execFile } from 'child_process';
import { classifyRunFailure, type ScriptRunner } from '../errors';
import { getConfig, reportRun, type RunBudget, timeoutFor } from '../runtime';

export interface ScriptProcess {
    runner: ScriptRunner;
    command: string;
    args: readonly string[];
    /** The script exactly as it runs, for errors and the debug hook. */
    script: string;
    budget: RunBudget;

    /** The script's own failure message, and the line it came from, out of stderr. */
    parseFailure(stderr: string): { message: string; line?: number };
}

export interface ProcessOutput {
    stdout: string;
    durationMs: number;
}

/** Run one script to completion, rejecting with the most specific OutlookError on failure. */
export function runScriptProcess(run: ScriptProcess): Promise<ProcessOutput> {
    const {maxBufferBytes, signal} = getConfig();
    const timeout = timeoutFor(run.budget);
    const startedAt = Date.now();
    return new Promise((resolve, reject) => {
        execFile(
            run.command,
            [...run.args],
            // windowsHide keeps a GUI host (an Electron app, say) from flashing a
            // console window on every run.
            {maxBuffer: maxBufferBytes, timeout, signal, windowsHide: true, encoding: 'utf8'},
            (error, stdout, stderr) => {
                const durationMs = Date.now() - startedAt;
                if (error) {
                    const {message, line} = run.parseFailure(stderr ?? '');
                    const failure = classifyRunFailure({
                        runner: run.runner,
                        script: run.script,
                        stderr: stderr ?? '',
                        message,
                        line,
                        durationMs,
                        nodeError: error,
                        timeoutMs: timeout,
                        maxBufferBytes,
                        signal,
                    });
                    reportRun({runner: run.runner, script: run.script, durationMs, error: failure.message});
                    reject(failure);
                    return;
                }
                reportRun({runner: run.runner, script: run.script, durationMs});
                resolve({stdout, durationMs});
            },
        );
    });
}
