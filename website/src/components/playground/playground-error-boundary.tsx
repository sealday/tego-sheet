import { Component, type ErrorInfo, type ReactNode } from 'react';
import styles from './playground.module.css';

interface PlaygroundErrorBoundaryProps {
  readonly children: ReactNode;
  readonly onReset: () => void;
}

interface PlaygroundErrorBoundaryState {
  readonly failed: boolean;
}

export class PlaygroundErrorBoundary extends Component<
  PlaygroundErrorBoundaryProps,
  PlaygroundErrorBoundaryState
> {
  public state: PlaygroundErrorBoundaryState = { failed: false };

  public static getDerivedStateFromError(): PlaygroundErrorBoundaryState {
    return { failed: true };
  }

  public componentDidCatch(_error: Error, _info: ErrorInfo): void {
    // Docusaurus reports the exception to its development overlay. The public recovery UI stays
    // deliberately terse and never renders an exception message or stack trace.
  }

  private readonly reset = (): void => {
    this.setState({ failed: false });
    this.props.onReset();
  };

  private readonly reload = (): void => {
    window.location.reload();
  };

  public render(): ReactNode {
    if (!this.state.failed) return this.props.children;

    return (
      <section className={styles.errorState} role="alert">
        <h2>The playground could not render this preset</h2>
        <p>Reset to the default preset or reload the page to start again.</p>
        <div className={styles.buttonRow}>
          <button type="button" onClick={this.reset}>
            Reset
          </button>
          <button type="button" onClick={this.reload}>
            Reload
          </button>
        </div>
      </section>
    );
  }
}
