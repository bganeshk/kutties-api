export interface Row {
  id: string;
  [key: string]: unknown;
}

export interface QueryOptions {
  filter?: Record<string, string>;
  limit?: number;
  offset?: number;
}

export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  total?: number;
}
