import { useEffect, useMemo, useState } from 'preact/hooks';
import { AuthError, latestOnly, type NasEntry, type NasRoot, type NasStore } from '../platform/nasStore.ts';
import { Dialog } from './Dialog.tsx';
import { Icon } from './Icon.tsx';

export type NasBrowserResult = { readonly root: string; readonly path: string };

/**
 * The NAS folder browser, for Open (pick a file) and Save As (pick a folder + name). Asks for the
 * server token when the server wants one.
 */
export function NasBrowser(props: {
  readonly store: NasStore;
  readonly mode: 'open' | 'save';
  readonly suggestedName?: string;
  readonly onDone: (result: NasBrowserResult | null) => void;
}) {
  const { store, mode } = props;
  const [roots, setRoots] = useState<NasRoot[] | null>(null);
  const [root, setRoot] = useState<string | null>(null);
  const [path, setPath] = useState('');
  const [entries, setEntries] = useState<NasEntry[]>([]);
  const [selected, setSelected] = useState<string | null>(null);
  const [name, setName] = useState(props.suggestedName ?? 'Untitled.ods');
  const [error, setError] = useState<string | null>(null);
  const [needsToken, setNeedsToken] = useState(false);
  const [token, setToken] = useState('');
  const listing = useMemo(() => latestOnly(), []);

  const fail = (e: unknown) => {
    if (e instanceof AuthError) setNeedsToken(true);
    else setError(e instanceof Error ? e.message : String(e));
  };

  const loadRoots = () =>
    store
      .roots()
      .then((r) => {
        setRoots(r);
        setNeedsToken(false);
        if (r.length === 1) setRoot(r[0]!.id);
      })
      .catch(fail);

  useEffect(() => {
    void loadRoots();
  }, []);

  useEffect(() => {
    if (!root) return;
    setSelected(null);
    listing(store.list(root, path), setEntries, fail);
  }, [root, path]);

  const crumbs = path ? path.split('/') : [];
  const open = (e: NasEntry) => {
    if (e.kind === 'dir') setPath(e.path);
    else if (mode === 'open') props.onDone({ root: root!, path: e.path });
    else setName(e.name);
  };
  const submit = () => {
    if (!root) return;
    if (mode === 'open') {
      const file = entries.find((e) => e.path === selected && e.kind === 'file');
      if (file) props.onDone({ root, path: file.path });
      return;
    }
    const clean = name.trim();
    if (!clean) return;
    const withExt = /\.(ods|csv|tsv)$/i.test(clean) ? clean : `${clean}.ods`;
    props.onDone({ root, path: path ? `${path}/${withExt}` : withExt });
  };

  if (needsToken)
    return (
      <Dialog
        title="NAS access token"
        onClose={() => props.onDone(null)}
        onSubmit={() => {
          store.token = token;
          void loadRoots();
        }}
        actions={
          <>
            <button type="button" class="cs-btn" onClick={() => props.onDone(null)}>
              Cancel
            </button>
            <button type="submit" class="cs-btn primary">
              Connect
            </button>
          </>
        }
      >
        <label for="cs-token">This server needs the token set in CELLESTRA_TOKEN.</label>
        <input
          id="cs-token"
          type="password"
          autofocus
          value={token}
          onInput={(e) => setToken((e.target as HTMLInputElement).value)}
        />
      </Dialog>
    );

  return (
    <Dialog
      title={mode === 'open' ? 'Open from NAS' : 'Save to NAS'}
      onClose={() => props.onDone(null)}
      onSubmit={submit}
      actions={
        <>
          <button type="button" class="cs-btn" onClick={() => props.onDone(null)}>
            Cancel
          </button>
          <button type="submit" class="cs-btn primary" disabled={mode === 'open' ? !selected : !root}>
            {mode === 'open' ? 'Open' : 'Save'}
          </button>
        </>
      }
    >
      {error && <p style={{ color: 'var(--cs-danger)' }}>{error}</p>}
      <div class="cs-crumbs">
        {roots && roots.length > 1 && (
          <select
            value={root ?? ''}
            onChange={(e) => {
              setRoot((e.target as HTMLSelectElement).value);
              setPath('');
            }}
            style={{ width: 'auto' }}
          >
            <option value="" disabled>
              Choose a folder…
            </option>
            {roots.map((r) => (
              <option key={r.id} value={r.id}>
                {r.name}
              </option>
            ))}
          </select>
        )}
        {roots && roots.length === 1 && (
          <button type="button" class="cs-btn" onClick={() => setPath('')}>
            {roots[0]!.name}
          </button>
        )}
        {crumbs.map((c, i) => (
          <span key={i}>
            <Icon name="chevron-right" />
            <button type="button" class="cs-btn" onClick={() => setPath(crumbs.slice(0, i + 1).join('/'))}>
              {c}
            </button>
          </span>
        ))}
      </div>
      {roots && roots.length === 0 && <p>The server has no folders configured (CELLESTRA_ROOTS).</p>}
      {root && (
        <ul class="cs-file-list" role="listbox" aria-label="Files">
          {path && (
            <li
              role="option"
              aria-selected={false}
              onDblClick={() => setPath(crumbs.slice(0, -1).join('/'))}
              onClick={() => setSelected(null)}
            >
              <Icon name="arrow-left" /> ..
            </li>
          )}
          {entries.map((e) => (
            <li
              key={e.path}
              role="option"
              aria-selected={selected === e.path}
              onClick={() => {
                setSelected(e.path);
                if (mode === 'save' && e.kind === 'file') setName(e.name);
              }}
              onDblClick={() => open(e)}
            >
              <Icon name={e.kind === 'dir' ? 'folder' : 'file'} />
              {e.name}
              {e.kind === 'file' && <span class="cs-meta">{new Date(e.modified).toLocaleString()}</span>}
            </li>
          ))}
          {entries.length === 0 && (
            <li role="option" aria-disabled="true" aria-selected={false} style={{ opacity: 0.6 }}>
              No spreadsheets here
            </li>
          )}
        </ul>
      )}
      {mode === 'save' && (
        <>
          <label for="cs-save-name">File name</label>
          <input
            id="cs-save-name"
            type="text"
            autofocus
            value={name}
            onInput={(e) => setName((e.target as HTMLInputElement).value)}
          />
        </>
      )}
    </Dialog>
  );
}
