import { DocumentViewer } from './DocumentViewer'

// A server component: the viewer itself is a client component.
export default function Page() {
  return (
    <main style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <DocumentViewer />
    </main>
  )
}
