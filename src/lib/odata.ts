// Escaping helpers for Microsoft Graph OData query parameters.

// OData string literals escape a single quote by doubling it. Applying this
// before URL-encoding means user-controlled text (e.g. a folder displayName)
// containing ' cannot break out of a $filter string literal (OData injection).
export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
