// Runs songslides.py inside Pyodide, off the main thread so the page stays responsive.
// Messages in:  {type: "generate", id, text, template (ArrayBuffer|null), toc, style}
//               {type: "word", id, document (ArrayBuffer)}         -> the .docx as "# Title" text
//               {type: "template", id, template (ArrayBuffer)}  -> the template's artwork as SVG
// Messages out: {type: "status", text} | {type: "ready"} | {type: "result", id, bytes, summary}
//               | {type: "template", id, art: {width_in, height_in, svg}} | {type: "word", id, text}
//               | {type: "error", id?, kind?, message}

const PYODIDE_VERSION = "0.29.5"; // the 314.x line fails to load in some current browsers
const PYODIDE_URL = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;
const PYTHON_PPTX = "python-pptx==1.0.2";
const VERSION = new URL(self.location.href).searchParams.get("v") || "dev";

importScripts(`${PYODIDE_URL}pyodide.js`);

const status = (text) => postMessage({ type: "status", text });

const ready = (async () => {
  status("Loading Python…");
  const pyodide = await loadPyodide({ indexURL: PYODIDE_URL });
  status("Loading the PowerPoint library…");
  await pyodide.loadPackage(["micropip", "lxml", "pillow", "typing-extensions"]);
  await pyodide.pyimport("micropip").install(PYTHON_PPTX);
  const source = await fetch(`songslides.py?v=${VERSION}`, { cache: "no-cache" }).then((r) => {
    if (!r.ok) throw new Error(`Could not load songslides.py (${r.status})`);
    return r.text();
  });
  pyodide.FS.writeFile("/home/pyodide/songslides.py", source);
  pyodide.runPython(`
import json, sys
sys.path.insert(0, "/home/pyodide")
import songslides

def run(text, template_path, toc, font_family, title_size, content_size):
    template = open(template_path, "rb").read() if template_path else None
    style = songslides.Style(font_family, title_size, content_size)
    data, summary = songslides.generate_pptx(text, template, toc=toc, style=style)
    with open("/tmp/out.pptx", "wb") as f:
        f.write(data)
    return json.dumps(summary.to_dict())

def word(path):
    return songslides.docx_to_text(open(path, "rb").read())

def preview(template_path):
    return json.dumps(songslides.template_preview(open(template_path, "rb").read()))
`);
  return pyodide;
})();

ready.then(
  () => postMessage({ type: "ready" }),
  (err) => postMessage({ type: "error", message: `The generator could not start: ${err.message}` }),
);

// Python tracebacks are long; the last line ("ValueError: No songs found…") is what people need.
function pythonMessage(err) {
  const lines = String(err.message || err).trim().split("\n");
  return lines[lines.length - 1].replace(/^\w+Error: /, "");
}

async function templateArt(pyodide, data) {
  try {
    pyodide.FS.writeFile("/tmp/preview.pptx", new Uint8Array(data.template));
    const preview = pyodide.globals.get("preview");
    const art = JSON.parse(preview("/tmp/preview.pptx"));
    preview.destroy();
    postMessage({ type: "template", id: data.id, art });
  } catch (err) {
    postMessage({ type: "error", id: data.id, kind: "template", message: pythonMessage(err) });
  }
}

async function wordText(pyodide, data) {
  try {
    pyodide.FS.writeFile("/tmp/songs.docx", new Uint8Array(data.document));
    const word = pyodide.globals.get("word");
    const text = word("/tmp/songs.docx");
    word.destroy();
    postMessage({ type: "word", id: data.id, text });
  } catch (err) {
    postMessage({ type: "error", id: data.id, kind: "word", message: pythonMessage(err) });
  }
}

onmessage = async ({ data }) => {
  const pyodide = await ready;
  if (data.type === "template") return templateArt(pyodide, data);
  if (data.type === "word") return wordText(pyodide, data);
  if (data.type !== "generate") return;
  try {
    let templatePath = null;
    if (data.template) {
      templatePath = "/tmp/template.pptx";
      pyodide.FS.writeFile(templatePath, new Uint8Array(data.template));
    }
    const { font_family, title_size, content_size } = data.style;
    const run = pyodide.globals.get("run");
    const summary = JSON.parse(run(data.text, templatePath, data.toc, font_family, title_size, content_size));
    run.destroy();
    const bytes = pyodide.FS.readFile("/tmp/out.pptx");
    postMessage({ type: "result", id: data.id, bytes, summary }, [bytes.buffer]);
  } catch (err) {
    postMessage({ type: "error", id: data.id, message: pythonMessage(err) });
  }
};
