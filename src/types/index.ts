export interface Row {
  id: string;
  [key: string]: unknown;
}

export interface QueryOptions {
  filter?: Record<string, string>;
  limit?: number;
  offset?: number;
  /** ISO date string (YYYY-MM-DD). Rows whose attendanceDate is before this are excluded. */
  sinceDate?: string;
}

export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  total?: number;
}
