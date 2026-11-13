export interface SidecarArtifactCommentRequest {
  uri?: string;
  comment?: string;
  selection?: string;
  pageTitle?: string;
}

export interface SidecarBridge {
  conversationId: string | null;
  fetch(url: string, options?: RequestInit): Promise<Response>;
  getWorkspaceUris(): Promise<string[]>;
  onWorkspaceChange(listener: (workspaceUris: string[]) => void): () => void;
  agent: {
    sendMessage(message: string, convId?: string): Promise<unknown>;
    startConversation(message: string, title?: string): Promise<unknown>;
    getConversationMetadata(conversationId?: string): Promise<unknown>;
  };
  ui: {
    toggleAuxPane(request?: Record<string, unknown>): void;
    toggleConversation(conversationId: string, options?: Record<string, unknown>): void;
    addArtifactComment(request?: SidecarArtifactCommentRequest): void;
  };
}

declare global {
  interface Window {
    sidecar?: SidecarBridge;
  }
}

export {};
