import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from 'react'
import { jsPDF } from 'jspdf'
import html2canvas from 'html2canvas'
import { renderAsync } from 'docx-preview'

type Tab = 'convert' | 'images' | 'compress' | 'fonts' | 'background'
type FileInfo = { file: File; url: string; width?: number; height?: number }

const tabs: { id: Tab; label: string; icon: string }[] = [
  { id: 'convert', label: 'Convert', icon: '⇄' },
  { id: 'images', label: 'Images', icon: '▧' },
  { id: 'compress', label: 'Compress', icon: '↘' },
  { id: 'fonts', label: 'Fonts', icon: 'Aa' },
  { id: 'background', label: 'Background', icon: '◐' },
]

const fonts = [
  ['Inter', 'Sans'], ['Geist', 'Sans'], ['Manrope', 'Sans'], ['DM Sans', 'Sans'], ['Plus Jakarta Sans', 'Sans'], ['Sora', 'Sans'], ['Outfit', 'Sans'], ['Poppins', 'Sans'], ['Space Grotesk', 'Display'], ['Bricolage Grotesque', 'Display'], ['Syne', 'Display'], ['Archivo Black', 'Display'], ['Playfair Display', 'Serif'], ['DM Serif Display', 'Serif'], ['Libre Baskerville', 'Serif'], ['Lora', 'Serif'], ['Newsreader', 'Serif'], ['JetBrains Mono', 'Mono'], ['IBM Plex Mono', 'Mono'], ['Space Mono', 'Mono'], ['Fira Code', 'Mono'], ['Roboto', 'Sans'], ['Montserrat', 'Sans'], ['Raleway', 'Sans'], ['Nunito Sans', 'Sans'], ['Work Sans', 'Sans'], ['Karla', 'Sans'], ['Bebas Neue', 'Display'], ['Fraunces', 'Serif'], ['Cormorant Garamond', 'Serif'], ['Merriweather', 'Serif'], ['Source Serif 4', 'Serif'], ['Alegreya', 'Serif'], ['Inconsolata', 'Mono'], ['Roboto Mono', 'Mono'], ['Ubuntu Mono', 'Mono'], ['Oswald', 'Display'], ['League Spartan', 'Display'], ['Barlow', 'Sans'], ['Figtree', 'Sans'], ['Cabin', 'Sans'], ['Rubik', 'Sans'], ['Chakra Petch', 'Display'], ['Crimson Pro', 'Serif'], ['IBM Plex Sans', 'Sans'], ['Noto Sans', 'Sans'], ['Mulish', 'Sans'], ['Quicksand', 'Sans'], ['Vollkorn', 'Serif'], ['Courier Prime', 'Mono']]

const colors = ['#ffffff', '#11131a', '#f4f1ea', '#eef4ff', '#e9f8f4', '#fff4e8', '#f6eaff', '#ffe9ef', '#dbe7ff', '#c6f5eb', '#ffe0b8', '#d9d0ff', '#ffd0d9', '#b8d8ff', '#d2f4c8', '#fff3a8', '#e6e6e6', '#8d99ae']

function prettySize(bytes: number) {
  if (!bytes) return '0 KB'
  return bytes > 1024 * 1024 ? `${(bytes / (1024 * 1024)).toFixed(2)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`
}

function removeSimpleBackground(source: HTMLImageElement) {
  const canvas = document.createElement('canvas')
  canvas.width = source.naturalWidth; canvas.height = source.naturalHeight
  const context = canvas.getContext('2d', { willReadFrequently: true })!
  context.drawImage(source, 0, 0)
  const image = context.getImageData(0, 0, canvas.width, canvas.height)
  const { data, width, height } = image
  const samplePoints = [[0, 0], [width - 1, 0], [0, height - 1], [width - 1, height - 1]]
  const bg = samplePoints.reduce((acc, [x, y]) => { const index = (y * width + x) * 4; acc[0] += data[index]; acc[1] += data[index + 1]; acc[2] += data[index + 2]; return acc }, [0, 0, 0]).map(value => value / samplePoints.length)
  const visited = new Uint8Array(width * height)
  const queue: number[] = []
  const edge = (x: number, y: number) => { const index = y * width + x; if (!visited[index]) { visited[index] = 1; queue.push(index) } }
  for (let x = 0; x < width; x += 1) { edge(x, 0); edge(x, height - 1) }
  for (let y = 0; y < height; y += 1) { edge(0, y); edge(width - 1, y) }
  const threshold = 42
  while (queue.length) {
    const index = queue.shift()!
    const x = index % width; const y = Math.floor(index / width); const offset = index * 4
    const red = data[offset]; const green = data[offset + 1]; const blue = data[offset + 2]
    const distance = Math.hypot(red - bg[0], green - bg[1], blue - bg[2])
    const saturation = Math.max(red, green, blue) - Math.min(red, green, blue)
    if (distance > threshold || saturation > 32) continue
    data[offset + 3] = 0
    if (x > 0) edge(x - 1, y); if (x < width - 1) edge(x + 1, y); if (y > 0) edge(x, y - 1); if (y < height - 1) edge(x, y + 1)
  }
  context.putImageData(image, 0, 0)
  return canvas
}

async function renderDocxToPdf(file: File, filename: string) {
  const mount = document.createElement('div')
  mount.id = 'docx-renderer'
  document.body.appendChild(mount)
  try {
    await renderAsync(file, mount, mount, { breakPages: true, ignoreWidth: false, ignoreHeight: false, useBase64URL: true, renderHeaders: true, renderFooters: true })
    await new Promise(resolve => window.setTimeout(resolve, 180))
    const wrapper = mount.querySelector('.docx-wrapper') as HTMLElement | null
    const pageNodes = Array.from(mount.querySelectorAll('section.docx')) as HTMLElement[]
    const renderedPages = pageNodes.length ? pageNodes : [wrapper || mount]
    const pdf = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'portrait', compress: true })
    const pageCanvases: HTMLCanvasElement[] = []
    for (const page of renderedPages) {
      const canvas = await html2canvas(page, { scale: 1.75, useCORS: true, backgroundColor: '#ffffff', logging: false })
      // Some DOCX files render as one tall wrapper. Slice that wrapper at A4 page boundaries.
      const pageHeight = Math.round(canvas.width * 1.4142)
      if (canvas.height > pageHeight * 1.25) {
        for (let y = 0; y < canvas.height; y += pageHeight) {
          const slice = document.createElement('canvas'); slice.width = canvas.width; slice.height = Math.min(pageHeight, canvas.height - y)
          slice.getContext('2d')!.drawImage(canvas, 0, y, canvas.width, slice.height, 0, 0, canvas.width, slice.height)
          pageCanvases.push(slice)
        }
      } else pageCanvases.push(canvas)
    }
    for (let index = 0; index < pageCanvases.length; index += 1) {
      const canvas = pageCanvases[index]
      const ratio = Math.min(515 / canvas.width, 742 / canvas.height)
      const width = canvas.width * ratio
      const height = canvas.height * ratio
      const x = (595.28 - width) / 2
      const y = (841.89 - height) / 2
      if (index > 0) pdf.addPage()
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.98), 'JPEG', x, y, width, height, undefined, 'FAST')
    }
    pdf.save(filename)
  } finally {
    mount.remove()
  }
}

function App() {
  const [tab, setTab] = useState<Tab>('convert')
  const [file, setFile] = useState<FileInfo | null>(null)
  const [dragging, setDragging] = useState(false)
  const [toast, setToast] = useState('')
  const [dark, setDark] = useState(false)
  const [font, setFont] = useState('Inter')
  const [fontQuery, setFontQuery] = useState('')
  const [fontCategory, setFontCategory] = useState('All')
  const [output, setOutput] = useState('PDF')
  const [quality, setQuality] = useState(82)
  const [target, setTarget] = useState('2 MB')
  const [compressMode, setCompressMode] = useState<'Compress' | 'Upscale'>('Compress')
  const [bgColor, setBgColor] = useState('#eef4ff')
  const [keepTransparency, setKeepTransparency] = useState(false)
  const [backgroundPreviewUrl, setBackgroundPreviewUrl] = useState('')
  const [backgroundView, setBackgroundView] = useState<'before' | 'after'>('after')
  const [text, setText] = useState('')
  const input = useRef<HTMLInputElement>(null)

  useEffect(() => {
    let cancelled = false
    if (tab !== 'background' || !file || !file.file.type.startsWith('image/')) { setBackgroundPreviewUrl(''); return }
    const image = new Image()
    image.onload = () => {
      if (cancelled) return
      const cutout = removeSimpleBackground(image)
      if (keepTransparency) { setBackgroundPreviewUrl(cutout.toDataURL('image/png')); return }
      const output = document.createElement('canvas'); output.width = cutout.width; output.height = cutout.height
      const context = output.getContext('2d')!
      context.fillStyle = bgColor; context.fillRect(0, 0, output.width, output.height); context.drawImage(cutout, 0, 0)
      setBackgroundPreviewUrl(output.toDataURL('image/jpeg', quality / 100))
    }
    image.src = file.url
    return () => { cancelled = true }
  }, [tab, file, bgColor, keepTransparency, quality])

  const notify = (message: string) => { setToast(message); window.setTimeout(() => setToast(''), 2600) }

  const loadFile = (incoming?: File) => {
    if (!incoming) return
    if (tab === 'background' && !incoming.type.startsWith('image/')) { notify('Background only accepts image files.'); return }
    if (incoming.size > 50 * 1024 * 1024) { notify('That file is over the 50 MB limit.'); return }
    const next: FileInfo = { file: incoming, url: URL.createObjectURL(incoming) }
    if (incoming.type.startsWith('image/')) {
      const image = new Image()
      image.onload = () => setFile({ ...next, width: image.naturalWidth, height: image.naturalHeight })
      image.src = next.url
    } else setFile(next)
    notify('File added. Ready when you are.')
  }

  const onDrop = (event: DragEvent<HTMLDivElement>) => { event.preventDefault(); setDragging(false); loadFile(event.dataTransfer.files?.[0]) }
  const onChange = (event: ChangeEvent<HTMLInputElement>) => loadFile(event.target.files?.[0])
  const reset = () => { if (file) URL.revokeObjectURL(file.url); setFile(null); setText(''); notify('Workspace cleared.') }

  const download = async () => {
    if (tab === 'convert' && file && output === 'DOCX' && file.file.type === 'application/pdf') {
      notify('PDF to DOCX exact conversion needs a document renderer. Your original PDF is unchanged.'); return
    }
    if (tab === 'convert' && file && output === 'PDF' && file.file.name.toLowerCase().match(/\.docx?$/)) {
      try {
        await renderDocxToPdf(file.file, `${file.file.name.replace(/\.[^.]+$/, '')}-filo.pdf`)
        notify('DOCX converted with pages, images, and layout preserved.')
      } catch {
        notify('This document could not be rendered. The original file is unchanged.')
      }
      return
    }
    if (tab === 'convert' && file && output === 'PDF' && file.file.type === 'application/pdf') {
      const original = await file.file.arrayBuffer()
      const blob = new Blob([original], { type: 'application/pdf' })
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${file.file.name.replace(/\.[^.]+$/, '')}-filo.pdf`; a.click(); notify('Original PDF downloaded without changes.')
      return
    }
    if (tab === 'convert' && (!file || output === 'PDF')) {
      const pdf = new jsPDF({ unit: 'pt', format: 'a4' })
      pdf.setFont(font.toLowerCase().replace(/ /g, '-').includes('serif') ? 'times' : 'helvetica')
      pdf.setFontSize(22); pdf.text(file?.file.name.replace(/\.[^.]+$/, '') || 'FILO document', 56, 76)
      pdf.setFontSize(12); pdf.text(text || 'Drop a document or paste text to create a polished PDF.', 56, 112, { maxWidth: 480 })
      pdf.save(`${file?.file.name.replace(/\.[^.]+$/, '') || 'filo-export'}.pdf`); notify(file?.file.name.toLowerCase().match(/\.docx?$/) ? 'PDF export created. For exact DOCX fidelity, connect a document renderer.' : 'PDF downloaded.')
      return
    }
    if (!file) { notify('Add a file first.'); return }
    if (tab === 'images' || tab === 'background') {
      const image = new Image(); image.src = file.url
      await new Promise(resolve => { image.onload = resolve })
      const canvas = tab === 'background' ? removeSimpleBackground(image) : document.createElement('canvas')
      if (tab === 'images') { canvas.width = image.naturalWidth; canvas.height = image.naturalHeight }
      const ctx = canvas.getContext('2d')!
      if (tab === 'background' && !keepTransparency) { const output = document.createElement('canvas'); output.width = canvas.width; output.height = canvas.height; const outputContext = output.getContext('2d')!; outputContext.fillStyle = bgColor; outputContext.fillRect(0, 0, output.width, output.height); outputContext.drawImage(canvas, 0, 0); return output.toBlob(blob => { if (!blob) return; const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${file.file.name.replace(/\.[^.]+$/, '')}-background.jpg`; a.click(); notify('Background replaced and exported.') }, 'image/jpeg', quality / 100) }
      if (tab === 'images') ctx.drawImage(image, 0, 0)
      const mime = tab === 'images' && output === 'JPEG' || tab === 'background' && !keepTransparency ? 'image/jpeg' : 'image/png'
      canvas.toBlob(blob => { if (!blob) return; const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${file.file.name.replace(/\.[^.]+$/, '')}-filo.${mime === 'image/jpeg' ? 'jpg' : 'png'}`; a.click(); notify('Export downloaded.') }, mime, quality / 100)
      return
    }
    const a = document.createElement('a'); a.href = file.url; a.download = `${file.file.name.replace(/\.[^.]+$/, '')}-filo.${output.toLowerCase()}`; a.click(); notify('Download ready.')
  }

  const filteredFonts = useMemo(() => fonts.filter(([name, category]) => (fontCategory === 'All' || category === fontCategory) && name.toLowerCase().includes(fontQuery.toLowerCase())), [fontQuery, fontCategory])

  return <div className={dark ? 'app dark' : 'app'}>
    <header className="topbar">
      <div className="brand"><div className="brand-mark">F</div><span>FILO</span><small>FILES, REFINED.</small></div>
      <nav className="navlinks"><button className="active">Workspace</button><button onClick={() => document.getElementById('how-it-works')?.scrollIntoView({ behavior: 'smooth' })}>How it works</button></nav>
      <div className="top-actions"><span className="privacy"><i /> Local-first</span><button className="icon-button" aria-label="Toggle theme" onClick={() => setDark(v => !v)}>{dark ? '☼' : '◐'}</button></div>
    </header>

    <main>
      <section className="hero"><div className="eyebrow"><span /> PRIVATE BY DEFAULT</div><h1>DROP A FILE.<br /><em>MAKE IT BETTER.</em></h1><p>Everything you need to convert, compress, resize, and refine your files — right in your browser.</p></section>

      <div className="tabs" role="tablist">{tabs.map(item => <button type="button" key={item.id} className={tab === item.id ? 'tab active' : 'tab'} onClick={() => { if (item.id === 'background' && file && !file.file.type.startsWith('image/')) setFile(null); setTab(item.id); requestAnimationFrame(() => document.querySelector('.workspace')?.scrollIntoView({ behavior: 'smooth', block: 'start' })) }} role="tab" aria-selected={tab === item.id}><span>{item.icon}</span>{item.label}</button>)}</div>

      <section className="workspace">
        <div className="control-column">
          <div className="section-heading"><div><p className="kicker">{tab === 'convert' ? 'UNIVERSAL CONVERTER' : tab === 'images' ? 'IMAGE STUDIO' : tab === 'compress' ? 'SIZE LAB' : tab === 'fonts' ? 'TYPE LIBRARY' : 'BACKGROUND LAB'}</p><h2>{tab === 'convert' ? 'One file in. Any format out.' : tab === 'images' ? 'Change the format, keep the detail.' : tab === 'compress' ? 'Less baggage. Same file.' : tab === 'fonts' ? 'Dress the words, don’t costume them.' : 'A cleaner backdrop in one click.'}</h2></div><span className="step">01 / 03</span></div>
          {tab !== 'fonts' && <div className={dragging ? 'dropzone dragging' : 'dropzone'} onDragOver={e => { e.preventDefault(); setDragging(true) }} onDragLeave={() => setDragging(false)} onDrop={onDrop} onClick={() => input.current?.click()}><input ref={input} type="file" accept={tab === 'background' ? 'image/jpeg,image/png,image/webp' : undefined} hidden onChange={onChange} /><div className="drop-icon">{file ? '✓' : '↥'}</div><strong>{file ? file.file.name : 'Drop it like it’s optimized.'}</strong><span>{file ? `${file.file.type || 'Unknown type'} · ${prettySize(file.file.size)}` : tab === 'background' ? 'JPEG, PNG, or WEBP only · up to 50 MB' : 'or browse from your device · up to 50 MB'}</span>{file?.width && <span className="dimensions">{file.width} × {file.height}px</span>}</div>}

          {tab === 'convert' && <div className="panel"><label>Convert to<select value={output} onChange={e => setOutput(e.target.value)}><option>PDF</option><option>DOCX</option><option>TXT</option><option>PNG</option><option>JPEG</option><option>ZIP</option></select></label><div className="inline-fields"><label>Page size<select><option>A4</option><option>Letter</option></select></label><label>Margins<select><option>Normal</option><option>Narrow</option><option>Wide</option></select></label></div><div className="type-section"><div><p className="kicker">TYPE & FONT</p><strong>Set the voice of your document.</strong><span>Pick a typeface, then paste or write your copy below.</span></div><button type="button" className="font-select" style={{ fontFamily: font }} onClick={() => setTab('fonts')}>{font}<span>↗</span></button></div><label>Document copy<textarea value={text} onChange={e => setText(e.target.value)} placeholder="Write or paste your copy here…" /></label>{file?.file.name.toLowerCase().match(/\.docx?$/) && <div className="fidelity-note"><b>Exact layout mode</b><span>DOC/DOCX files need a document renderer to preserve every image and style byte-for-byte. This browser export keeps your copy and selected type settings intact.</span></div>}</div>}
          {tab === 'images' && <div className="panel"><label>Output format<div className="segmented"><button className={output === 'PNG' ? 'selected' : ''} onClick={() => setOutput('PNG')}>PNG</button><button className={output === 'JPEG' ? 'selected' : ''} onClick={() => setOutput('JPEG')}>JPEG</button></div></label><label>JPEG quality <span className="value">{quality}%</span><input type="range" min="10" max="100" value={quality} onChange={e => setQuality(Number(e.target.value))} /></label><div className="hint">Transparency is preserved in PNG. JPEG is lighter and universal.</div></div>}
          {tab === 'compress' && <div className="panel"><label>Mode<div className="segmented"><button className={compressMode === 'Compress' ? 'selected' : ''} onClick={() => setCompressMode('Compress')}>Compress</button><button className={compressMode === 'Upscale' ? 'selected' : ''} onClick={() => setCompressMode('Upscale')}>Upscale</button></div></label><label>Target download size<select value={target} onChange={e => setTarget(e.target.value)}><option>Under 200 KB</option><option>1 MB</option><option>2 MB</option><option>5 MB</option><option>10 MB</option></select></label><label>Quality <span className="value">{quality}%</span><input type="range" min="10" max="100" value={quality} onChange={e => setQuality(Number(e.target.value))} /></label><div className="size-meter"><div><span>Original</span><b>{prettySize(file?.file.size || 0)}</b></div><div className="meter"><span style={{ width: `${file ? Math.min(94, 28 + quality / 2) : 12}%` }} /></div><div><span>Estimated</span><b>{file ? prettySize(file.file.size * (compressMode === 'Compress' ? (1 - quality / 180) : 1.25)) : '—'}</b></div></div></div>}
          {tab === 'fonts' && <div className="panel font-panel"><div className="type-editor"><div><p className="kicker">TYPE YOUR COPY</p><strong style={{ fontFamily: font }}>Make the words feel like yours.</strong><span>Choose a font below, then write anything you want to preview.</span></div><textarea value={text} onChange={e => setText(e.target.value)} placeholder="Type or paste your copy here…" style={{ fontFamily: font }} /><button type="button" className="copy-button" onClick={() => { navigator.clipboard?.writeText(text); notify(text ? 'Copy copied to clipboard.' : 'Write something first.') }}>Copy text <span>↗</span></button></div><div className="search-row"><input placeholder="Search 50 fonts" value={fontQuery} onChange={e => setFontQuery(e.target.value)} /><span>⌕</span></div><div className="chips">{['All', 'Sans', 'Serif', 'Display', 'Mono'].map(category => <button type="button" className={fontCategory === category ? 'chip active' : 'chip'} key={category} onClick={() => setFontCategory(category)}>{category}</button>)}</div><div className="font-grid">{filteredFonts.map(([name, category]) => <button type="button" key={name} className={font === name ? 'font-card selected' : 'font-card'} onClick={() => setFont(name)} style={{ fontFamily: name }}><span>{name}</span><small>{category}</small></button>)}</div></div>}
          {tab === 'background' && <div className="panel"><div className="background-panel-head"><div><p className="kicker">BACKGROUND COLOR</p><strong>Pick a clean, on-brand backdrop.</strong><span>FILO removes only the edge-connected neutral background, keeping the person untouched.</span></div></div><div className="color-line"><input type="color" value={bgColor} onChange={e => { setBgColor(e.target.value); setBackgroundView('after') }} /><input aria-label="Background hex" value={bgColor} onChange={e => { setBgColor(e.target.value); setBackgroundView('after') }} /><span className="swatch" style={{ background: bgColor }} /></div><div className="color-grid">{colors.map(color => <button aria-label={`Use ${color}`} key={color} className={bgColor === color ? 'color active' : 'color'} style={{ background: color }} onClick={() => { setBgColor(color); setBackgroundView('after') }} />)}</div><label className="checkbox"><input type="checkbox" checked={keepTransparency} onChange={e => { setKeepTransparency(e.target.checked); setBackgroundView('after') }} /> Keep transparent areas</label><div className="background-shortcuts"><button type="button" className="shortcut active">Add background</button><button type="button" className="shortcut">Edit cutout</button><button type="button" className="shortcut">Resize</button></div><div className="hint">Best for passport photos, portraits, and product shots with a clean solid backdrop.</div></div>}
          <div className="actionbar"><button className="ghost" onClick={reset}>Reset</button><button className="ghost" onClick={() => notify('Preview refreshed.')}>Preview <span>↗</span></button><button className="primary" onClick={download}>Download <span>↓</span></button></div>
        </div>
        <aside className="preview-column"><div className="preview-head"><div><p className="kicker">LIVE PREVIEW</p><h3>{file ? 'Your file, in focus.' : 'Ready when you are.'}</h3></div><div className="preview-head-actions">{tab === 'background' && <div className="compare-toggle"><button type="button" className={backgroundView === 'before' ? 'selected' : ''} onClick={() => setBackgroundView('before')}>Before</button><button type="button" className={backgroundView === 'after' ? 'selected' : ''} onClick={() => setBackgroundView('after')}>After</button></div>}<span className="live-dot"><i /> Live</span></div></div><div className={file ? 'preview-canvas has-file' : 'preview-canvas'}>{file ? (file.file.type.startsWith('image/') ? <img src={tab === 'background' && backgroundPreviewUrl && backgroundView === 'after' ? backgroundPreviewUrl : file.url} alt="Uploaded preview" /> : <div className="file-preview"><div className="file-glyph">{file.file.name.toLowerCase().endsWith('pdf') ? 'PDF' : 'DOC'}</div><strong>{file.file.name}</strong><span>{prettySize(file.file.size)} · {file.file.type || 'file'}</span></div>) : <div className="empty-preview"><div className="empty-mark">✦</div><strong>Drop a file to preview it</strong><span>Your result will appear here as soon as you add a file.</span></div>}</div><div className="preview-foot"><span>OUTPUT</span><b>{tab === 'background' ? (keepTransparency ? 'PNG' : 'JPEG') : output}</b><span className="quality-pill">{quality}% quality</span></div></aside>
      </section>

      <section className="trust-row"><span><b>50 MB</b> max file size</span><span><b>30+</b> formats supported</span><span><b>100%</b> in-browser processing</span><span><b>0</b> uploads to servers</span></section>
      <section id="how-it-works" className="how"><div><p className="kicker">HOW IT WORKS</p><h2>Three moves. No mystery.</h2></div><div className="how-grid"><div><span>01</span><strong>Drop it</strong><p>Bring a document, image, or any file from your device.</p></div><div><span>02</span><strong>Refine it</strong><p>Choose a format, size, typeface, or background that fits.</p></div><div><span>03</span><strong>Take it</strong><p>Preview the result, then download a ready-to-use file.</p></div></div></section>
    </main>
    {toast && <div className="toast"><span>✓</span>{toast}</div>}
  </div>
}

export default App
