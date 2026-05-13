"use client";

import * as React from "react";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";

type ResultErrorBoundaryState = {
  error: Error | null;
};

export class ResultErrorBoundary extends React.Component<React.PropsWithChildren, ResultErrorBoundaryState> {
  state: ResultErrorBoundaryState = {
    error: null
  };

  static getDerivedStateFromError(error: Error): ResultErrorBoundaryState {
    return { error };
  }

  render() {
    if (this.state.error) {
      return (
        <Alert variant="destructive">
          <AlertTitle>Result panel failed to render</AlertTitle>
          <AlertDescription>{this.state.error.message}</AlertDescription>
        </Alert>
      );
    }

    return this.props.children;
  }
}
