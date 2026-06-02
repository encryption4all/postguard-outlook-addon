// DOM helpers shared by taskpane views.

// Resolve an element by id and throw if it's missing. The taskpane views
// reference their controls by id and would crash on a typo anyway —
// `byId` makes the failure loud and types the result so call sites
// don't need to narrow.
export function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el as T;
}
