const svgs = import.meta.glob<string>('../icons/*.svg', { query: '?raw', import: 'default', eager: true });
const byName = new Map(
  Object.entries(svgs).map(([path, svg]) => [path.replace(/^.*\/(.+)\.svg$/, '$1'), svg]),
);

/** A Lucide icon by name (`bold`, `save`, …). Inline SVG so it inherits `currentColor`. */
export function Icon({ name, title }: { readonly name: string; readonly title?: string }) {
  const svg = byName.get(name);
  if (!svg) return null;
  return (
    <span
      class="cs-icon"
      aria-hidden={title ? undefined : true}
      title={title}
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  );
}
