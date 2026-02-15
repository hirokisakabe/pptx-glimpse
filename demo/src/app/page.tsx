import { UploadViewer } from "@/components/UploadViewer";

export default function Home() {
  return (
    <div className="app">
      <header>
        <h1>pptx-glimpse</h1>
        <p>Upload a PPTX file and preview slides as SVG</p>
      </header>
      <main>
        <UploadViewer />
      </main>
      <footer>
        <a href="https://github.com/hirokisakabe/pptx-glimpse">GitHub</a>
        <span> | </span>
        <a href="https://www.npmjs.com/package/pptx-glimpse">npm</a>
      </footer>
    </div>
  );
}
