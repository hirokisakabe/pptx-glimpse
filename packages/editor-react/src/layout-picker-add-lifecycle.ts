export interface LayoutPickerAddCompletion {
  readonly generation: number;
  readonly owner: object;
}

/** Prevents deferred add completions from affecting a replacement editor session. */
export class LayoutPickerAddLifecycle {
  private generation = 0;
  private owner: object;
  private active = true;

  constructor(owner: object) {
    this.owner = owner;
  }

  activate(owner: object): void {
    if (owner !== this.owner) this.generation += 1;
    this.owner = owner;
    this.active = true;
  }

  capture(): LayoutPickerAddCompletion {
    return { generation: this.generation, owner: this.owner };
  }

  isCurrent(completion: LayoutPickerAddCompletion): boolean {
    return (
      this.active && completion.generation === this.generation && completion.owner === this.owner
    );
  }

  invalidate(): void {
    this.generation += 1;
    this.active = false;
  }
}
