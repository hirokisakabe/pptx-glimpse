interface CompositionCompletion {
  readonly promise: Promise<boolean>;
  readonly resolve: (completed: boolean) => void;
}

/** Tracks transient direct-editor work without allowing one editor session to affect another. */
export class DirectTextEditorLifecycle {
  private generation = 0;
  private compositionCompletion: CompositionCompletion | null = null;
  private commit: Promise<boolean> | null = null;

  currentGeneration(): number {
    return this.generation;
  }

  isCurrent(generation: number): boolean {
    return generation === this.generation;
  }

  beginComposition(): void {
    this.cancelComposition();
    let resolveComposition: (completed: boolean) => void = () => {};
    const promise = new Promise<boolean>((resolve) => {
      resolveComposition = resolve;
    });
    this.compositionCompletion = {
      promise,
      resolve: resolveComposition,
    };
  }

  completeComposition(): void {
    const completion = this.compositionCompletion;
    this.compositionCompletion = null;
    completion?.resolve(true);
  }

  cancelComposition(): void {
    this.compositionCompletion?.resolve(false);
    this.compositionCompletion = null;
  }

  compositionPromise(): Promise<boolean> | undefined {
    return this.compositionCompletion?.promise;
  }

  currentCommit(): Promise<boolean> | null {
    return this.commit;
  }

  setCommit(generation: number, commit: Promise<boolean>): boolean {
    if (!this.isCurrent(generation) || this.commit !== null) return false;
    this.commit = commit;
    return true;
  }

  clearCommit(generation: number): void {
    if (this.isCurrent(generation)) this.commit = null;
  }

  invalidate(): void {
    this.generation += 1;
    this.cancelComposition();
    this.commit = null;
  }
}
