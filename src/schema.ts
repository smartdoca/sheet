/** Current provisioning default. Recovery always uses the version in its atomic bundle. */
export const EXLSX_SCHEMA_VERSION = 6 as const
export const EXLSX_SHEET_STATE_SCHEMA_VERSION = 2 as const
export const EXLSX_FEATURE_SCHEMA_VERSION = 3 as const
export type ExlsxSchemaVersion = 1 | 2 | 3 | 4 | 5 | 6
