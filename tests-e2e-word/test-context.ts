/**
 * E2E Test Context (office-coding-agent pattern)
 *
 * Singleton chứa kết quả test được gửi từ Word task pane về Node test runner.
 * Tránh giới hạn URL query ~8KB bằng cách POST JSON qua fetch.
 */

export interface WordTestResult {
  name: string;
  pass: boolean;
  durationMs: number;
  error?: string;
  details?: Record<string, unknown>;
}

class E2ETestContext {
  private _results: WordTestResult[] = [];

  setResults(results: WordTestResult[]) {
    this._results = results;
  }

  getResult(name: string): WordTestResult | undefined {
    return this._results.find((r) => r.name === name);
  }

  getAllResults(): WordTestResult[] {
    return this._results;
  }

  getFailed(): WordTestResult[] {
    return this._results.filter((r) => !r.pass);
  }

  getPassed(): WordTestResult[] {
    return this._results.filter((r) => r.pass);
  }
}

export const e2eContext = new E2ETestContext();
