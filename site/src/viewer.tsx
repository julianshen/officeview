import { forwardRef, useMemo, type CSSProperties, type Ref } from 'react'
import { OfficeFile, type OfficeFileProps, type OfficeViewerHandle } from 'officeview'
import { useSiteTheme } from './layout'

/** The viewer worker for large spreadsheets, bundled from the library source. */
export const createViewerWorker = () => new Worker(new URL('../../src/worker/worker-entry.ts', import.meta.url), { type: 'module' })

type ViewerProps = Omit<OfficeFileProps, 'data'> & {
  /** A URL to fetch, or bytes/a Blob already in hand. */
  src: string | Uint8Array | Blob | null
  height?: CSSProperties['height']
  frameStyle?: CSSProperties
}

/**
 * <OfficeFile> in a sized frame (the viewer tracks pages by scrolling its own
 * container, so it needs a constrained height), with a progress bar and a
 * readable error, following the site theme unless one is passed.
 */
export const Viewer = forwardRef(function Viewer({ src, height = 640, frameStyle, theme, ...props }: ViewerProps, ref: Ref<OfficeViewerHandle>) {
  const [siteTheme] = useSiteTheme()
  // A stable source object: a new one would re-download the file.
  const data = useMemo(() => (typeof src === 'string' ? { url: src } : src), [src])
  return (
    <div className="viewer-frame" style={{ height, ...frameStyle }}>
      {data === null ? (
        <div className="viewer-state">No document</div>
      ) : (
        <OfficeFile
          ref={ref}
          data={data}
          theme={theme ?? siteTheme}
          viewerWorker={createViewerWorker}
          loading={({ loaded, total }) => (
            <div className="viewer-state" role="status">
              Loading document{total ? ` · ${Math.round((loaded / total) * 100)}%` : '…'}
              <div className="bar"><span style={{ width: total ? `${(loaded / total) * 100}%` : '30%' }} /></div>
            </div>
          )}
          error={(message) => <div className="viewer-state error" role="alert">Could not open this file: {message}</div>}
          {...props}
        />
      )}
    </div>
  )
})
