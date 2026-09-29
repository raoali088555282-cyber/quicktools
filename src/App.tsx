import { useEffect, useMemo, useRef, useState } from 'react';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { jsPDF } from 'jspdf';
import { PDFDocument } from 'pdf-lib';
import * as pdfjsLib from 'pdfjs-dist';
import QRCode from 'qrcode';
import {
  Download,
  FileText,
  Image as ImageIcon,
  Menu,
  QrCode,
  Receipt,
  Search,
  Sparkles,
  X,
  ArrowRight,
  Check,
  Shield,
  Zap,
  FileOutput,
  ScanLine,
  BriefcaseBusiness,
  Quote,
} from 'lucide-react';


// Vercel-compatible local QR API.
// This replaces the AppDeploy-only client so the frontend can build anywhere.
// QR records are stored in this browser's localStorage.
type LocalQrRecord = {
  id: string;
  target: string;
  scans: { at: string; device: string }[];
};

const QR_STORAGE_KEY = 'quicktools_qr_records';

function readQrRecords(): LocalQrRecord[] {
  try {
    const raw = localStorage.getItem(QR_STORAGE_KEY);
    return raw ? (JSON.parse(raw) as LocalQrRecord[]) : [];
  } catch {
    return [];
  }
}

function writeQrRecords(records: LocalQrRecord[]) {
  localStorage.setItem(QR_STORAGE_KEY, JSON.stringify(records));
}

function getDeviceLabel(): string {
  const ua = navigator.userAgent;
  if (/iPhone|iPad|iPod/i.test(ua)) return 'iPhone/iPad';
  if (/Android/i.test(ua)) return 'Android';
  if (/Windows/i.test(ua)) return 'Windows';
  if (/Mac/i.test(ua)) return 'Mac';
  if (/Linux/i.test(ua)) return 'Linux';
  return 'Unknown';
}

const api = {
  async post(path: string, body: { target: string }) {
    if (path !== '/api/qr') throw new Error('Unsupported API route');
    const id = crypto.randomUUID().replace(/-/g, '').slice(0, 12);
    const record: LocalQrRecord = { id, target: body.target, scans: [] };
    const records = readQrRecords();
    records.push(record);
    writeQrRecords(records);
    return { data: { id: record.id, target: record.target } };
  },

  async get(path: string) {
    const analyticsMatch = path.match(/^\/api\/qr\/([^/]+)\/analytics$/);
    if (analyticsMatch) {
      const record = readQrRecords().find(item => item.id === analyticsMatch[1]);
      return { data: { scans: record?.scans || [] } };
    }

    const redirectMatch = path.match(/^\/api\/qr\/([^/]+)\/redirect$/);
    if (redirectMatch) {
      const records = readQrRecords();
      const index = records.findIndex(item => item.id === redirectMatch[1]);
      if (index === -1) throw new Error('NOT_FOUND');

      records[index].scans.push({
        at: new Date().toISOString(),
        device: getDeviceLabel(),
      });
      writeQrRecords(records);

      return { data: { target: records[index].target } };
    }

    throw new Error('Unsupported API route');
  },
};

type ToolKey =
  | 'pdf-to-word'
  | 'pdf-compressor'
  | 'background-remover'
  | 'resume-builder'
  | 'qr-code-generator'
  | 'qr-analytics'
  | 'invoice-generator'
  | 'quotation-generator';

const tools = [
  {
    key: 'pdf-to-word',
    title: 'PDF to Word',
    description: 'Turn PDF text into an editable DOCX file.',
    icon: FileOutput,
    category: 'PDF Tools',
  },
  {
    key: 'pdf-compressor',
    title: 'PDF Compressor',
    description: 'Reduce PDF file size and download the result.',
    icon: FileText,
    category: 'PDF Tools',
  },
  {
    key: 'background-remover',
    title: 'Background Remover',
    description: 'Create a transparent PNG from a simple image background.',
    icon: ImageIcon,
    category: 'Image Tools',
  },
  {
    key: 'resume-builder',
    title: 'Resume Builder',
    description: 'Build a polished resume and export it as PDF.',
    icon: BriefcaseBusiness,
    category: 'Resume',
  },
  {
    key: 'qr-code-generator',
    title: 'QR Code Generator',
    description: 'Create downloadable QR codes for any valid URL.',
    icon: QrCode,
    category: 'QR Tools',
  },
  {
    key: 'qr-analytics',
    title: 'QR Analytics',
    description: 'Create trackable QR links and view real scan events.',
    icon: ScanLine,
    category: 'QR Tools',
  },
  {
    key: 'invoice-generator',
    title: 'Invoice Generator',
    description: 'Create professional invoices with automatic totals.',
    icon: Receipt,
    category: 'Business Tools',
  },
  {
    key: 'quotation-generator',
    title: 'Quotation Generator',
    description: 'Create polished quotations with automatic totals.',
    icon: Quote,
    category: 'Business Tools',
  },
] as const;

const seo: Record<
  string,
  { title: string; description: string; h1: string; keywords: string }
> = {
  home: {
    title: 'QuickTools — Free Online Tools. Fast Results.',
    description:
      'Use free online PDF, image, resume, QR, invoice and quotation tools with fast, simple workflows.',
    h1: 'Free Online Tools. Fast Results.',
    keywords:
      'free online tools, PDF tools, resume builder, QR code generator, invoice generator',
  },
  'pdf-to-word': {
    title: 'PDF to Word Converter Online – Free PDF to DOCX | QuickTools',
    description:
      'Convert PDF to Word online for free. Turn PDF files into editable DOCX documents with a simple browser-based tool.',
    h1: 'PDF to Word Converter',
    keywords: 'PDF to Word converter, PDF to DOCX, convert PDF to Word online, free PDF converter',
  },
  'pdf-compressor': {
    title: 'Compress PDF Online Free | QuickTools',
    description:
      'Compress PDF files online for free while reducing file size quickly and easily.',
    h1: 'PDF Compressor',
    keywords:
      'PDF compressor, compress PDF, reduce PDF size, compress PDF online',
  },
  'background-remover': {
    title: 'Image Background Remover Online – Free | QuickTools',
    description:
      'Remove backgrounds from images online for free and create transparent PNG images with a simple browser-based background remover.',
    h1: 'Image Background Remover',
    keywords: 'background remover, remove background from image, transparent PNG, image background remover online',
  },
  'resume-builder': {
    title: 'Free Resume Builder Online – Create & Download PDF | QuickTools',
    description:
      'Create a professional resume online for free, add your experience and skills, then download a clean PDF resume.',
    h1: 'Resume Builder',
    keywords: 'free resume builder, CV maker, resume maker, create resume online, resume PDF',
  },
  'qr-code-generator': {
    title: 'Free QR Code Generator Online – Create QR Codes | QuickTools',
    description:
      'Create QR codes online for free from website URLs. Generate downloadable QR code images quickly in your browser.',
    h1: 'QR Code Generator',
    keywords: 'free QR code generator, QR code maker, URL QR code, create QR code online',
  },
  'qr-analytics': {
    title: 'QR Code Tracking & Analytics – Track Scans Online | QuickTools',
    description:
      'Create trackable QR links and view scan events with simple QR code analytics and tracking in your browser.',
    h1: 'QR Analytics',
    keywords: 'QR code tracking, QR analytics, trackable QR code, QR scan tracking',
  },
  'invoice-generator': {
    title: 'Free Invoice Generator Online – Create PDF Invoices | QuickTools',
    description:
      'Create professional invoices online for free, calculate totals, and download a printable PDF invoice in seconds.',
    h1: 'Invoice Generator',
    keywords:
      'free invoice generator, invoice maker, create invoice online, invoice PDF',
  },
  'quotation-generator': {
    title: 'Free Quotation Generator Online – Create PDF Quotes | QuickTools',
    description:
      'Create professional quotations online for free, calculate totals, and download a clear PDF quote for customers.',
    h1: 'Quotation Generator',
    keywords:
      'free quotation generator, quote generator, quotation maker, quotation PDF',
  },
};

function getRoute() {
  const raw = window.location.pathname.replace(/^\/+|\/+$/g, '');
  if (!raw) return 'home';
  if (raw === 'tools') return 'tools';
  const match = raw.match(
    /^tools\/(pdf-to-word|pdf-compressor|background-remover|resume-builder|qr-code-generator|qr-analytics|invoice-generator|quotation-generator)$/
  );
  if (match) return raw;
  const scanMatch = raw.match(/^scan\/([A-Za-z0-9_-]+)$/);
  if (scanMatch) return raw;
  return raw === 'privacy' || raw === 'terms' || raw === 'contact' ? raw : 'home';
}

function navigate(route: string) {
  const depth = window.location.pathname.split('/').filter(Boolean).length;
  const prefix = '../'.repeat(depth);
  const target = route === 'home' ? prefix || './' : `${prefix}${route}`;
  window.history.pushState({}, '', target);
  window.dispatchEvent(new PopStateEvent('popstate'));
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function StructuredData({ route }: { route: string }) {
  const baseUrl = window.location.origin;
  const toolKey = route.startsWith('tools/') ? route.slice('tools/'.length) : '';
  const toolSeo = seo[toolKey];
  const data = route.startsWith('tools/') && toolSeo
    ? [
        {
          '@context': 'https://schema.org',
          '@type': 'WebApplication',
          name: toolSeo.h1,
          description: toolSeo.description,
          url: `${baseUrl}/tools/${toolKey}`,
          applicationCategory: 'UtilitiesApplication',
          operatingSystem: 'Web',
          offers: { '@type': 'Offer', price: '0', priceCurrency: 'USD' },
        },
        {
          '@context': 'https://schema.org',
          '@type': 'BreadcrumbList',
          itemListElement: [
            { '@type': 'ListItem', position: 1, name: 'QuickTools', item: `${baseUrl}/` },
            { '@type': 'ListItem', position: 2, name: 'Tools', item: `${baseUrl}/tools` },
            { '@type': 'ListItem', position: 3, name: toolSeo.h1, item: `${baseUrl}/tools/${toolKey}` },
          ],
        },
      ]
    : {
        '@context': 'https://schema.org',
        '@type': 'WebSite',
        name: 'QuickTools',
        url: `${baseUrl}/`,
        description: seo.home.description,
      };
  return (
    <>
      {Array.isArray(data) ? data.map((item, index) => (
        <script key={index} type="application/ld+json" dangerouslySetInnerHTML={{ __html: JSON.stringify(item) }} />
      )) : (
        <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: JSON.stringify(data) }} />
      )}
    </>
  );
}
function AdSlot({ position }: { position: string }) {
  useEffect(() => {
    try {
      const ads = (window as typeof window & { adsbygoogle?: unknown[] }).adsbygoogle;
      if (ads) ads.push({});
    } catch {
      // AdSense may be unavailable until the site is approved or ads are ready.
    }
  }, []);

  return (
    <div className="ad-slot" aria-label={`Advertisement ${position}`}>
      <ins
        className="adsbygoogle"
        style={{ display: 'block' }}
        data-ad-client="ca-pub-3492290120117899"
        data-ad-slot="4919929348"
        data-ad-format="auto"
        data-full-width-responsive="true"
      />
    </div>
  );
}

function Layout({
  children,
  route,
  onNavigate,
}: {
  children: React.ReactNode;
  route: string;
  onNavigate: (r: string) => void;
}) {
  const [mobile, setMobile] = useState(false);
  const [query, setQuery] = useState('');
  const results = useMemo(
    () =>
      tools
        .filter(t =>
          `${t.title} ${t.description}`
            .toLowerCase()
            .includes(query.toLowerCase())
        )
        .slice(0, 5),
    [query]
  );
  return (
    <div className="site-shell">
      <header className="topbar">
        <button
          className="brand"
          onClick={() => onNavigate('home')}
          aria-label="QuickTools home"
        >
          <span className="brand-mark">Q</span>
          <span>QuickTools</span>
        </button>
        <nav className={mobile ? 'nav open' : 'nav'}>
          <button onClick={() => onNavigate('tools')}>Tools</button>
          <button onClick={() => onNavigate('tools')}>PDF Tools</button>
          <button onClick={() => onNavigate('tools/resume-builder')}>
            Resume
          </button>
          <button onClick={() => onNavigate('tools/qr-code-generator')}>
            QR Tools
          </button>
          <button onClick={() => onNavigate('tools/invoice-generator')}>
            Business Tools
          </button>
        </nav>
        <div className="nav-search">
          <Search size={17} />
          <input
            value={query}
            onChange={e => setQuery(e.target.value)}
            placeholder="Search tools..."
            aria-label="Search tools"
          />
          {query && (
            <div className="search-results">
              {results.map(t => (
                <button
                  key={t.key}
                  onClick={() => {
                    onNavigate(`tools/${t.key}`);
                    setQuery('');
                  }}
                >
                  {t.title}
                  <ArrowRight size={14} />
                </button>
              ))}
            </div>
          )}
        </div>
        <button
          className="mobile-menu"
          onClick={() => setMobile(!mobile)}
          aria-label="Menu"
        >
          {mobile ? <X /> : <Menu />}
        </button>
      </header>
      <main>{children}</main>
      <footer className="footer">
        <div>
          <button className="footer-brand" onClick={() => onNavigate('home')}>
            QuickTools
          </button>
          <p>Free Tools. Fast Results.</p>
        </div>
        <div className="footer-links">
          <button onClick={() => onNavigate('tools')}>All Tools</button>
          <button onClick={() => onNavigate('privacy')}>Privacy</button>
          <button onClick={() => onNavigate('terms')}>Terms</button>
          <button onClick={() => onNavigate('contact')}>Contact</button>
        </div>
        <div className="footer-copy">
          © {new Date().getFullYear()} QuickTools
        </div>
      </footer>
    </div>
  );
}

function ToolCard({
  tool,
  onOpen,
}: {
  tool: (typeof tools)[number];
  onOpen: () => void;
}) {
  const Icon = tool.icon;
  return (
    <button className="tool-card" onClick={onOpen}>
      <span className="tool-icon">
        <Icon size={23} />
      </span>
      <span className="tool-card-copy">
        <strong>{tool.title}</strong>
        <small>{tool.description}</small>
      </span>
      <ArrowRight className="tool-arrow" size={18} />
    </button>
  );
}

function Home({ onNavigate }: { onNavigate: (r: string) => void }) {
  return (
    <>
      <section className="hero">
        <div className="hero-glow" />
        <div className="eyebrow">
          <Sparkles size={15} /> SIMPLE TOOLS, SERIOUSLY USEFUL
        </div>
        <h1>
          Free Online Tools.
          <br />
          <span>Fast Results.</span>
        </h1>
        <p>
          Convert, compress, create and manage your files with simple, powerful
          online tools.
        </p>
        <div className="hero-actions">
          <button className="primary" onClick={() => onNavigate('tools')}>
            Explore All Tools <ArrowRight size={18} />
          </button>
          <button
            className="secondary"
            onClick={() => onNavigate('tools/pdf-compressor')}
          >
            Try PDF Compressor
          </button>
        </div>
        <div className="trust-row">
          <span>
            <Zap size={16} /> Fast workflows
          </span>
          <span>
            <Shield size={16} /> Privacy-minded
          </span>
          <span>
            <Check size={16} /> No fake results
          </span>
        </div>
      </section>
      <AdSlot position="homepage-top" />
      <section className="section">
        <div className="section-head">
          <div>
            <span className="eyebrow">EXPLORE</span>
            <h2>Tools for everyday work</h2>
          </div>
          <button className="text-button" onClick={() => onNavigate('tools')}>
            View all <ArrowRight size={16} />
          </button>
        </div>
        <div className="tool-grid">
          {tools.slice(0, 6).map(t => (
            <ToolCard
              key={t.key}
              tool={t}
              onOpen={() => onNavigate(`tools/${t.key}`)}
            />
          ))}
        </div>
      </section>
      <section className="feature-band">
        <div>
          <span className="eyebrow">BUILT AROUND THE WORKFLOW</span>
          <h2>One task at a time, without the clutter.</h2>
          <p>
            QuickTools keeps each utility focused, with dedicated pages, clear
            actions and helpful next steps.
          </p>
        </div>
        <div className="workflow">
          <div>
            <b>01</b>
            <span>Upload or enter</span>
          </div>
          <div>
            <b>02</b>
            <span>Process or create</span>          </div>
          <div>
            <b>03</b>
            <span>Download and continue</span>
          </div>
        </div>
      </section>
      <AdSlot position="homepage-mid" />
      <section className="section">
        <div className="section-head">
          <div>
            <span className="eyebrow">WHY QUICKTOOLS</span>
            <h2>Designed to be useful first.</h2>
          </div>
        </div>
        <div className="benefit-grid">
          <article>
            <Zap />
            <h3>Focused tools</h3>
            <p>
              Each utility has its own workspace instead of burying everything
              on one page.
            </p>
          </article>
          <article>
            <Shield />
            <h3>Honest results</h3>
            <p>
              No fabricated conversions, analytics, file sizes or success
              messages.
            </p>
          </article>
          <article>
            <Sparkles />
            <h3>Easy discovery</h3>
            <p>
              Related tools and clear navigation help you move naturally from
              one task to the next.
            </p>
          </article>
        </div>
      </section>
      <AdSlot position="homepage-bottom" />
    </>
  );
}

function ToolsDirectory({ onNavigate }: { onNavigate: (r: string) => void }) {
  return (
    <PageShell
      title="All Online Tools"
      intro="Explore QuickTools utilities for PDFs, images, resumes, QR codes and business documents."
      seoKey="home"
    >
      <div className="tool-grid">
        {tools.map(t => (
          <ToolCard
            key={t.key}
            tool={t}
            onOpen={() => onNavigate(`tools/${t.key}`)}
          />
        ))}
      </div>
    </PageShell>
  );
}

function PageShell({
  title,
  intro,
  seoKey,
  children,
}: {
  title: string;
  intro: string;
  seoKey: string;
  children: React.ReactNode;
}) {
  return (
    <>
      <section className="page-hero">
        <div className="eyebrow">QUICKTOOLS</div>
        <h1>{title}</h1>
        <p>{intro}</p>
      </section>
      <AdSlot position="top" />
      <div className="page-content">{children}</div>
    </>
  );
}

function ToolInterface({ toolKey }: { toolKey: ToolKey }) {
  if (toolKey === 'pdf-to-word') return <PDFToWord />;
  if (toolKey === 'pdf-compressor') return <PDFCompressor />;
  if (toolKey === 'background-remover') return <BackgroundRemover />;
  if (toolKey === 'resume-builder') return <ResumeBuilder />;
  if (toolKey === 'qr-code-generator')
    return <QRGenerator onNavigate={() => navigate('tools/qr-analytics')} />;
  if (toolKey === 'qr-analytics') return <QRAnalytics />;
  if (toolKey === 'invoice-generator') return <BusinessDoc mode="invoice" />;
  return <BusinessDoc mode="quotation" />;
}

function ToolPage({
  toolKey,
  onNavigate,
}: {
  toolKey: ToolKey;
  onNavigate: (r: string) => void;
}) {
  const item = tools.find(t => t.key === toolKey)!;
  return (
    <PageShell
      title={seo[toolKey].h1}
      intro={seo[toolKey].description}
      seoKey={toolKey}
    >
      <div className="tool-workspace">
        <ToolInterface toolKey={toolKey} />
      </div>
      <AdSlot position="result" />
      <section className="content-card">
        {toolKey === 'pdf-compressor' ? (
          <>
            <h2>Free PDF Compressor Online</h2>
            <p>
              Compress PDF files online with QuickTools to reduce file size for easier sharing, uploading and storage. This free PDF compressor is designed for a simple browser-based workflow: choose a PDF, process it, review the result and download the compressed file.
            </p>
            <h2>How to compress a PDF</h2>
            <ol>
              <li>Select a PDF file from your device or drag it into the upload area.</li>
              <li>Let QuickTools process the PDF and calculate the resulting file size.</li>
              <li>Review the result and download the compressed PDF.</li>
            </ol>
            <h2>Why compress a PDF?</h2>
            <p>
              Smaller PDF files are easier to attach to email, upload to websites, send through messaging apps and keep in cloud storage. Compression can also help when a website has a strict file-size limit.
            </p>
            <h2>PDF compression without complicated software</h2>
            <p>
              QuickTools runs in your web browser, so you can use the PDF compressor without installing a desktop application. The interface is focused on one task and provides the actual processed result instead of a simulated completion message.
            </p>
            <h2>Frequently asked questions</h2>
            <details>
              <summary>Is the PDF compressor free?</summary>
              <p>Yes. The core PDF compression tool is available for free. Advertising may be used to support QuickTools.</p>
            </details>
            <details>
              <summary>Will compression always make a PDF much smaller?</summary>
              <p>Not necessarily. The amount of size reduction depends on the PDF's contents, images, fonts and existing compression. QuickTools reports the result produced by the actual processing.</p>
            </details>
            <details>
              <summary>Can I compress a PDF on my phone?</summary>
              <p>Yes. QuickTools is designed to work in modern web browsers on phones, tablets and computers.</p>
            </details>
            <details>
              <summary>Do I need to install a PDF compressor?</summary>
              <p>No. You can use the browser-based QuickTools PDF compressor without installing desktop software.</p>
            </details>
            <details>
              <summary>Is my PDF stored?</summary>
              <p>Most file-processing tools on QuickTools are designed to process files in the browser. Review the Privacy Policy for current data-handling details and exceptions.</p>
            </details>
            <h2>More free PDF tools</h2>
            <p>
              After compressing a PDF, you can also use QuickTools <button className="inline-link" onClick={() => onNavigate('tools/pdf-to-word')}>PDF to Word Converter</button> to create an editable DOCX document.
            </p>
          </>
        ) : toolKey === 'pdf-to-word' ? (
          <>
            <h2>Free PDF to Word Converter Online</h2>
            <p>Convert PDF to Word online for free with QuickTools. Turn readable PDF text into an editable DOCX document directly in your browser, without installing desktop conversion software.</p>
            <h2>How to convert PDF to Word</h2>
            <ol>
              <li>Select a PDF file from your device.</li>
              <li>Let QuickTools extract readable text and create a Word-compatible DOCX file.</li>
              <li>Download the DOCX file and continue editing it in Microsoft Word or another compatible editor.</li>
            </ol>
            <h2>Why convert a PDF to an editable Word document?</h2>
            <p>Converting PDF to DOCX can make it easier to update text, reuse information, add notes, correct wording or continue working on an existing document. The final layout can vary because PDF files are designed for fixed-page presentation rather than Word-style editing.</p>
            <h2>What types of PDFs work best?</h2>
            <p>Text-based PDFs with selectable, readable text are generally better suited to this converter. Scanned documents, image-only PDFs, unusual fonts and complex multi-column layouts may require additional editing after conversion.</p>
            <h2>PDF to Word on desktop or mobile</h2>
            <p>QuickTools is browser-based, so you can use the PDF to Word converter on supported desktop, tablet and mobile browsers. The basic conversion workflow does not require installing a separate PDF-to-DOCX application.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is this PDF to Word converter free?</summary><p>Yes. The core PDF to Word conversion tool is available for free. Advertising may be used to support QuickTools.</p></details>
            <details><summary>Can I convert every PDF perfectly?</summary><p>No. Conversion quality depends on the PDF's text, fonts, structure and layout. Scanned or image-only PDFs may not convert as expected.</p></details>
            <details><summary>Does the converter create a DOCX file?</summary><p>Yes. QuickTools generates a Word-compatible DOCX document from readable PDF text.</p></details>
            <details><summary>Can I convert a PDF to Word on my phone?</summary><p>Yes. You can use the tool from a modern mobile browser without installing a separate conversion app.</p></details>
            <details><summary>Will the original PDF formatting always stay the same?</summary><p>No. PDF and Word use different document structures, so complex formatting may need manual adjustment after conversion.</p></details>
            <h2>More free PDF tools</h2>
            <p>Need to reduce a document's file size instead? Try the <button className="inline-link" onClick={() => onNavigate('tools/pdf-compressor')}>Free PDF Compressor</button>.</p>
          </>
        ) : toolKey === 'background-remover' ? (
          <>
            <h2>Free Image Background Remover Online</h2>
            <p>Remove a simple image background online and create a transparent PNG with QuickTools. The tool is designed for quick browser-based image editing without complicated software.</p>
            <h2>How to remove an image background</h2>
            <ol>
              <li>Choose an image from your device.</li>
              <li>Let QuickTools process the image and detect the background.</li>
              <li>Review the result and download the transparent PNG.</li>
            </ol>
            <h2>When to use a background remover</h2>
            <p>A transparent background can be useful for product images, profile graphics, presentations, thumbnails and simple design projects. Results depend on the contrast and complexity of the original image.</p>
            <h2>Transparent PNG without desktop software</h2>
            <p>Use the background remover directly in your browser on supported devices. No separate desktop image editor is required for the basic workflow.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is the background remover free?</summary><p>Yes. The core background-removal tool is available for free.</p></details>
            <details><summary>Does it work with every image?</summary><p>Results vary with image quality, background complexity, edges and subject contrast. Simple backgrounds generally provide clearer results.</p></details>
            <details><summary>What format can I download?</summary><p>The tool is designed to produce a transparent PNG so the removed background can remain transparent.</p></details>
            <details><summary>Can I remove a background on mobile?</summary><p>Yes. Use a modern mobile browser to access the tool.</p></details>
          </>
        ) : toolKey === 'resume-builder' ? (
          <>
            <h2>Free Resume Builder Online</h2>
            <p>Create a professional resume online with QuickTools. Add your contact details, profile, experience, education and skills, preview the document and export it as a PDF.</p>
            <h2>How to create a resume online</h2>
            <ol>
              <li>Enter your name, professional title and contact details.</li>
              <li>Add a concise profile summary, experience, education and relevant skills.</li>
              <li>Review the live preview and export your resume as a PDF.</li>
            </ol>
            <h2>What to include in a resume</h2>
            <p>A useful resume normally focuses on relevant experience, measurable achievements, education and skills that match the role. Keep information clear and easy to scan, and review your contact details before exporting.</p>
            <h2>Resume maker for students and professionals</h2>
            <p>QuickTools provides a simple browser-based workflow for creating a resume without starting from a blank document in a desktop editor.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is the resume builder free?</summary><p>Yes. The core resume builder is available for free.</p></details>
            <details><summary>Can I download my resume as a PDF?</summary><p>Yes. The builder includes an Export PDF option for downloading your completed resume.</p></details>
            <details><summary>What sections can I add?</summary><p>You can add a professional title, contact details, profile summary, experience, education and skills.</p></details>
            <details><summary>Can I use it on a phone?</summary><p>Yes. The builder is designed to work in modern web browsers on mobile and desktop devices.</p></details>
          </>
        ) : toolKey === 'qr-code-generator' ? (
          <>
            <h2>Free QR Code Generator Online</h2>
            <p>Generate a QR code for a website URL with QuickTools and download it as PNG or SVG. The generator is designed for fast, simple QR creation directly in your browser.</p>
            <h2>How to create a QR code</h2>
            <ol>
              <li>Enter a valid website URL beginning with http:// or https://.</li>
              <li>Click Generate QR Code.</li>
              <li>Download the QR code as a PNG or SVG file.</li>
            </ol>
            <h2>Where can you use a QR code?</h2>
            <p>QR codes can be used on menus, posters, business cards, packaging, presentations and printed materials to help people open a website quickly with a compatible camera or QR scanner.</p>
            <h2>Static QR code vs trackable QR code</h2>
            <p>This generator creates a QR image for the URL you enter. If you need QuickTools tracking for scan events, use the <button className="inline-link" onClick={() => onNavigate('tools/qr-analytics')}>QR Analytics</button> tool instead.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is the QR code generator free?</summary><p>Yes. The QR code generator is available for free.</p></details>
            <details><summary>Can I download the QR code?</summary><p>Yes. You can download generated QR codes as PNG or SVG.</p></details>
            <details><summary>What URLs can I use?</summary><p>Use a valid http:// or https:// website URL.</p></details>
            <details><summary>Can I track scans?</summary><p>Use QuickTools QR Analytics to create a trackable QR link and view scan events recorded by the browser-based system.</p></details>
          </>
        ) : toolKey === 'qr-analytics' ? (
          <>
            <h2>QR Code Analytics and Tracking</h2>
            <p>Create a trackable QR link with QuickTools and view scan events recorded by the browser-based analytics system. This page is useful when you need more than a static QR image.</p>
            <h2>How QR tracking works</h2>
            <ol>
              <li>Enter a valid destination URL.</li>
              <li>Create a trackable QR link.</li>
              <li>Download the generated trackable QR code and share it.</li>
              <li>Review recorded scan events in the analytics panel.</li>
            </ol>
            <h2>What QR analytics can show</h2>
            <p>QuickTools records scan events with a timestamp and a broad device label such as Android, iPhone/iPad, Windows, Mac or Linux. It does not claim to provide a complete marketing analytics platform.</p>
            <h2>Important tracking limitation</h2>
            <p>QR Analytics in QuickTools uses browser local storage for its records. Data is therefore tied to the browser/device where the records were created rather than a centralized cloud analytics account.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is QR Analytics free?</summary><p>Yes. The current QR analytics workflow is available for free.</p></details>
            <details><summary>Does it track every scan worldwide?</summary><p>No. The current implementation stores records locally in the browser, so it is not a centralized worldwide scan database.</p></details>
            <details><summary>What device information is recorded?</summary><p>The system records a broad device label and the scan timestamp for recorded events.</p></details>
            <details><summary>Can I download a trackable QR code?</summary><p>Yes. After creating a trackable link, you can download its QR code.</p></details>
          </>
        ) : toolKey === 'invoice-generator' ? (
          <>
            <h2>Free Invoice Generator Online</h2>
            <p>Create a simple professional invoice online with QuickTools. Add your business name, customer, items, quantities, prices, tax and discount, then download the invoice as a PDF.</p>
            <h2>How to make an invoice</h2>
            <ol>
              <li>Enter your business name and customer information.</li>
              <li>Add products or services with quantities and prices.</li>
              <li>Enter optional tax and discount percentages.</li>
              <li>Review the calculated total and download the PDF invoice.</li>
            </ol>
            <h2>Invoice calculations made simple</h2>
            <p>QuickTools calculates the subtotal from your line items, applies the tax percentage and subtracts the discount percentage to produce the displayed total.</p>
            <h2>Who can use an online invoice maker?</h2>
            <p>Freelancers, small businesses, service providers and independent sellers can use a simple invoice workflow when they need a quick document for a customer.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is the invoice generator free?</summary><p>Yes. The core invoice generator is available for free.</p></details>
            <details><summary>Can I download the invoice as a PDF?</summary><p>Yes. The invoice can be exported as a PDF from the browser.</p></details>
            <details><summary>Can I add tax and discounts?</summary><p>Yes. You can enter tax and discount percentages and the total updates automatically.</p></details>
            <details><summary>Do I need to install accounting software?</summary><p>No. The basic invoice workflow runs in your browser. It is a document-generation tool, not a replacement for full accounting software.</p></details>
          </>
        ) : toolKey === 'quotation-generator' ? (
          <>
            <h2>Free Quotation Generator Online</h2>
            <p>Create a professional quotation online with QuickTools. Add your business and customer details, line items, quantities, prices, tax and discount, then download a PDF quotation.</p>
            <h2>How to create a quotation</h2>
            <ol>
              <li>Enter your business name and customer information.</li>
              <li>Add the products or services you want to quote.</li>
              <li>Set quantities, prices, tax and discount values.</li>
              <li>Review the total and download the quotation as a PDF.</li>
            </ol>
            <h2>Why use an online quote generator?</h2>
            <p>A quotation gives a customer a clear summary of proposed products or services and their prices before a purchase or project begins. QuickTools keeps the basic creation process simple and browser-based.</p>
            <h2>Quotation maker for small businesses</h2>
            <p>The tool can be useful for freelancers, agencies, contractors and small businesses that need a quick quote document without opening a full accounting application.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is the quotation generator free?</summary><p>Yes. The core quotation generator is available for free.</p></details>
            <details><summary>Can I download a quotation PDF?</summary><p>Yes. You can export the quotation as a PDF from the browser.</p></details>
            <details><summary>Can I add tax and discounts?</summary><p>Yes. Tax and discount percentages can be entered and included in the displayed total.</p></details>
            <details><summary>Is this a full accounting system?</summary><p>No. QuickTools provides a focused quotation document workflow rather than a complete accounting or invoicing platform.</p></details>
          </>
        ) : (
          <>
            <h2>How to use {item.title}</h2>
            <ol>
              <li>Open the tool and provide the required file, link or information.</li>
              <li>Follow the on-screen processing or editing steps.</li>
              <li>Review the real result, then download or save it.</li>
            </ol>
            <h2>Why use QuickTools?</h2>
            <p>QuickTools focuses on clear, practical workflows. The page only reports results that the browser or service actually produces, and errors are shown instead of hidden.</p>
            <h2>Frequently asked questions</h2>
            <details><summary>Is this tool free?</summary><p>Yes. QuickTools provides the core utility for free. Advertising may be used to support the service.</p></details>
            <details><summary>Are my results real?</summary><p>Yes. Results shown by the tool come from the actual browser processing performed on your file or input.</p></details>
            <details><summary>Do I need to install software?</summary><p>No. QuickTools runs in your web browser, so you can use the tool without installing desktop software.</p></details>
            <details><summary>Is my uploaded file stored?</summary><p>Most file-processing tools run directly in your browser. Check the Privacy Policy for details about data handling and any exceptions.</p></details>
          </>
        )}
      </section>
      <AdSlot position="bottom" />
      <section className="related">
        <div className="section-head">
          <div>
            <span className="eyebrow">KEEP GOING</span>
            <h2>Related tools</h2>
          </div>
        </div>
        <div className="tool-grid">
          {tools
            .filter(t => t.key !== toolKey)
            .slice(0, 3)
            .map(t => (
              <ToolCard
                key={t.key}
                tool={t}
                onOpen={() => onNavigate(`tools/${t.key}`)}
              />
            ))}
        </div>
      </section>
    </PageShell>
  );
}

function FileDrop({
  accept,
  onFile,
}: {
  accept: string;
  onFile: (file: File) => void;
}) {
  const input = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  return (
    <div
      className={drag ? 'dropzone drag' : 'dropzone'}
      onDragOver={e => {
        e.preventDefault();
        setDrag(true);
      }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => {
        e.preventDefault();
        setDrag(false);
        const f = e.dataTransfer.files[0];
        if (f) onFile(f);
      }}
      onClick={() => input.current?.click()}
    >
      <input
          ref={input}
          hidden
          type="file"
          accept={accept}
          onClick={e => e.stopPropagation()}
          onChange={e => e.target.files?.[0] && onFile(e.target.files[0])}
        />
      <div className="upload-icon">
        <Download size={25} />
      </div>
      <strong>Drop a file here or click to browse</strong>
      <span>Supported: {accept}</span>
    </div>
  );
}

function PDFToWord() {
  const [file, setFile] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState('');
  async function convert() {
    if (!file) {
      setMsg('Please choose a PDF first.');
      return;
    }
    setBusy(true);
    setMsg('Reading PDF text…');
    try {
      const buffer = await file.arrayBuffer();
      const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(buffer), disableWorker: true });
      const pdf = await loadingTask.promise;
      const paragraphs: Paragraph[] = [];
      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
        const page = await pdf.getPage(pageNumber);
        const content = await page.getTextContent();
        let line = '';
        let lastY: number | null = null;
        for (const item of content.items) {
          if (!('str' in item)) continue;
          const text = item.str;
          const y = item.transform[5];
          if (lastY !== null && Math.abs(y - lastY) > 4) {
            if (line.trim()) paragraphs.push(new Paragraph({ children: [new TextRun(line.trim())] }));
            line = '';
          }
          if (text) line += (line ? ' ' : '') + text;
          lastY = y;
        }
        if (line.trim()) paragraphs.push(new Paragraph({ children: [new TextRun(line.trim())] }));
        if (pageNumber < pdf.numPages) paragraphs.push(new Paragraph({ children: [new TextRun('')] }));
      }
      if (!paragraphs.length) throw new Error('NO_TEXT');
      const doc = new Document({ sections: [{ properties: {}, children: paragraphs }] });
      const blob = await Packer.toBlob(doc);
      saveBlob(blob, file.name.replace(/\.pdf$/i, '') + '.docx');
      setMsg('DOCX created successfully from the PDF text.');
    } catch (error) {
      console.error(error);
      setMsg('No readable text was found. This tool currently requires a text-based PDF; scanned PDFs need OCR.');
    } finally {
      setBusy(false);
    }
  }
  return (
    <div className="tool-panel">
      <FileDrop accept=".pdf,application/pdf" onFile={setFile} />
      {file && (
        <div className="file-row">
          <FileText size={19} />
          <span>{file.name}</span>
          <b>{formatBytes(file.size)}</b>
        </div>
      )}
      <button
        className="primary full"
        disabled={!file || busy}
        onClick={convert}
      >
        {busy ? 'Creating DOCX…' : 'Convert to Word'} <ArrowRight size={17} />
      </button>
      {msg && (
        <div
          className={
            msg.includes('success') ? 'status success' : 'status error'
          }
        >
          {msg}
        </div>
      )}
    </div>
  );
}

function PDFCompressor() {
  const [file, setFile] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState('');
  async function compress() {
    if (!file) return;
    setBusy(true);
    setResult('');
    try {
      const src = await PDFDocument.load(await file.arrayBuffer());
      const out = await PDFDocument.create();
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach(p => out.addPage(p));
      const bytes = await out.save({ useObjectStreams: true });
      const blob = new Blob([bytes], { type: 'application/pdf' });
      const reduction = Math.max(
        0,
        ((file.size - blob.size) / file.size) * 100
      );
      saveBlob(blob, file.name.replace(/\.pdf$/i, '') + '-compressed.pdf');
      setResult(
        `Done — original ${formatBytes(file.size)}, result ${formatBytes(blob.size)}, actual reduction ${reduction.toFixed(1)}%.`
      );
    } catch {
      setResult(
        'The PDF could not be processed. Please try another valid PDF.'
      );
    } finally {
      setBusy(false);
    }
  }
  return (
    <div className="tool-panel">
      <FileDrop accept=".pdf,application/pdf" onFile={setFile} />
      {file && (
        <div className="file-row">
          <FileText size={19} />
          <span>{file.name}</span>
          <b>{formatBytes(file.size)}</b>
        </div>
      )}
      <button
        className="primary full"
        disabled={!file || busy}
        onClick={compress}
      >
        {busy ? 'Compressing…' : 'Compress PDF'} <ArrowRight size={17} />
      </button>
      {result && (
        <div
          className={
            result.startsWith('Done') ? 'status success' : 'status error'
          }
        >
          {result}
        </div>
      )}
    </div>
  );
}

function BackgroundRemover() {
  const [src, setSrc] = useState<string>('');
  const [out, setOut] = useState<string>('');
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState('');
  function process(file: File) {
    setBusy(true);
    setMsg('');
    const reader = new FileReader();
    reader.onload = () => {
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = image.width;
        canvas.height = image.height;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          setMsg('Canvas processing is unavailable.');
          setBusy(false);
          return;
        }
        ctx.drawImage(image, 0, 0);
        const data = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const d = data.data;
        const edge = Math.max(
          10,
          Math.floor(Math.min(canvas.width, canvas.height) * 0.08)
        );
        for (let y = 0; y < canvas.height; y++)
          for (let x = 0; x < canvas.width; x++) {
            const i = (y * canvas.width + x) * 4;
            const nearEdge =
              x < edge ||
              y < edge ||
              x >= canvas.width - edge ||
              y >= canvas.height - edge;
            const r = d[i],
              g = d[i + 1],
              b = d[i + 2];
            const bright = r > 220 && g > 220 && b > 220;
            const nearNeutral = Math.max(r, g, b) - Math.min(r, g, b) < 18;
            if (nearEdge && (bright || nearNeutral)) d[i + 3] = 0;
          }        ctx.putImageData(data, 0, 0);
        setSrc(reader.result as string);
        setOut(canvas.toDataURL('image/png'));
        setMsg(
          'Processed using a simple edge/background heuristic. Complex backgrounds may need a dedicated AI model.'
        );
        setBusy(false);
      };
      image.src = reader.result as string;
    };
    reader.readAsDataURL(file);
  }
  return (
    <div className="tool-panel">
      <FileDrop
        accept="image/png,image/jpeg,.png,.jpg,.jpeg"
        onFile={process}
      />
      {busy && <div className="progress">Processing image…</div>}
      {out && (
        <div className="compare">
          <div>
            <span>Original</span>
            <img src={src} alt="Original upload" />
          </div>
          <div>
            <span>Result</span>
            <img src={out} alt="Background removed result" />
          </div>
        </div>
      )}
      {out && (
        <button
          className="primary full"
          onClick={() => saveDataUrl(out, 'quicktools-background-removed.png')}
        >
          <Download size={17} /> Download transparent PNG
        </button>
      )}
      {msg && <div className="status">{msg}</div>}
    </div>
  );
}

type Resume = {
  name: string;
  title: string;
  email: string;
  phone: string;
  summary: string;
  skills: string;
  education: string;
  experience: string;
};
const blankResume: Resume = {
  name: '',
  title: '',
  email: '',
  phone: '',
  summary: '',
  skills: '',
  education: '',
  experience: '',
};

function ResumeBuilder() {
  const [r, setR] = useState<Resume>(() => {
    try {
      return JSON.parse(localStorage.getItem('qt-resume') || 'null') || blankResume;
    } catch {
      return blankResume;
    }
  });
  function update(k: keyof Resume, v: string) {
    const next = { ...r, [k]: v };
    setR(next);
    localStorage.setItem('qt-resume', JSON.stringify(next));
  }
  function exportPdf() {
    const doc = new jsPDF();
    let y = 20;
    const add = (text: string, size = 11, bold = false) => {
      doc.setFontSize(size);
      doc.setFont('helvetica', bold ? 'bold' : 'normal');
      const lines = doc.splitTextToSize(text, 175);
      doc.text(lines, 18, y);
      y += lines.length * (size * 0.48) + 6;
      if (y > 275) {
        doc.addPage();
        y = 20;
      }
    };
    add(r.name || 'Your Name', 22, true);
    add(r.title || 'Professional Title', 13, true);
    add([r.email, r.phone].filter(Boolean).join(' • '), 10);
    if (r.summary) {
      add('PROFILE', 12, true);
      add(r.summary);
    }
    if (r.experience) {
      add('EXPERIENCE', 12, true);
      add(r.experience);
    }
    if (r.education) {
      add('EDUCATION', 12, true);
      add(r.education);
    }
    if (r.skills) {
      add('SKILLS', 12, true);
      add(r.skills);
    }
    doc.save('quicktools-resume.pdf');
  }
  const fields: [keyof Resume, string, string][] = [
    ['name', 'Full name', 'Alex Morgan'],
    ['title', 'Professional title', 'Product Designer'],
    ['email', 'Email', 'alex@example.com'],
    ['phone', 'Phone', '+1 000 000 0000'],
    ['summary', 'Profile summary', 'A concise professional summary...'],
    ['experience', 'Experience', 'Company — Role — Dates — Achievements'],
    ['education', 'Education', 'Degree — Institution — Year'],
    ['skills', 'Skills', 'Design, research, communication'],
  ];
  return (
    <div className="resume-grid">
      <div className="form-panel">
        {fields.map(([k, l, p]) => (
          <label key={k}>
            {l}
            {k === 'summary' ||
            k === 'experience' ||
            k === 'education' ||
            k === 'skills' ? (
              <textarea
                value={r[k]}
                onChange={e => update(k, e.target.value)}
                placeholder={p}
                rows={k === 'summary' ? 4 : 3}
              />
            ) : (
              <input
                value={r[k]}
                onChange={e => update(k, e.target.value)}
                placeholder={p}
              />
            )}
          </label>
        ))}
        <button className="primary full" onClick={exportPdf}>
          <Download size={17} /> Export PDF
        </button>
      </div>
      <div className="resume-preview">
        <div className="resume-paper">
          <h2>{r.name || 'Your Name'}</h2>
          <h3>{r.title || 'Professional Title'}</h3>
          <p>{[r.email, r.phone].filter(Boolean).join(' • ')}</p>
          {[
            ['PROFILE', r.summary],
            ['EXPERIENCE', r.experience],
            ['EDUCATION', r.education],
            ['SKILLS', r.skills],
          ].map(([h, v]) =>
            v ? (
              <section key={h as string}>
                <b>{h}</b>
                <p>{v}</p>
              </section>
            ) : null
          )}
        </div>
      </div>
    </div>
  );
}

function QRGenerator({ onNavigate }: { onNavigate: (r: string) => void }) {
  const [url, setUrl] = useState('');
  const [img, setImg] = useState('');
  const [errorMsg, setErrorMsg] = useState('');
  async function make() {
    setErrorMsg('');
    try {
      const u = new URL(url);
      if (!/^https?:$/.test(u.protocol)) throw new Error();
      const data = await QRCode.toDataURL(url, { width: 640, margin: 2 });
      setImg(data);
    } catch {
      setErrorMsg('Enter a valid http:// or https:// URL.');
    }
  }
  return (
    <div className="qr-panel">
      <div className="qr-form">
        <label>
          Website URL
          <input
            value={url}
            onChange={e => setUrl(e.target.value)}
            placeholder="https://example.com"
          />
        </label>
        <button className="primary full" onClick={make}>
          Generate QR Code <QrCode size={17} />
        </button>
        {errorMsg && <div className="status error">{errorMsg}</div>}
        <p className="helper">
          Want real scan tracking? Use{' '}
          <button
            className="inline-link"
            onClick={() => onNavigate('tools/qr-analytics')}
          >
            QR Analytics
          </button>
          .
        </p>
      </div>
      {img && (
        <div className="qr-result">
          <img src={img} alt="Generated QR code" />
          <div>
            <button
              className="secondary"
              onClick={() => saveDataUrl(img, 'quicktools-qr.png')}
            >
              <Download size={16} /> PNG
            </button>
            <button className="secondary" onClick={() => downloadSvg(url)}>
              <Download size={16} /> SVG
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function QRAnalytics() {
  const [url, setUrl] = useState('');
  const [created, setCreated] = useState<{ id: string; target: string } | null>(null);
  const [scans, setScans] = useState<{ at: string; device: string }[]>([]);
  const [errorMsg, setErrorMsg] = useState('');
  const [busy, setBusy] = useState(false);
  async function create() {
    setErrorMsg('');
    try {
      const u = new URL(url);
      if (!/^https?:$/.test(u.protocol)) throw new Error();
      setBusy(true);
      const response = await api.post('/api/qr', { target: u.toString() });
      const record = response.data as { id: string; target: string };
      setCreated(record);
      setScans([]);
    } catch {
      setCreated(null);
      setScans([]);
      setErrorMsg('Enter a valid http:// or https:// URL, and try again.');
    } finally {
      setBusy(false);
    }
  }
  useEffect(() => {
    if (!created) return;
    let cancelled = false;
    const load = async () => {
      try {
        const response = await api.get(`/api/qr/${created.id}/analytics`);
        if (!cancelled) setScans((response.data?.scans || []) as { at: string; device: string }[]);
      } catch {
        if (!cancelled) setErrorMsg('Analytics could not be loaded right now.');
      }
    };
    void load();
    const timer = window.setInterval(load, 15000);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, [created]);
  return (
    <div className="tool-panel">
      <label>
        Destination URL
        <input
          value={url}
          onChange={e => setUrl(e.target.value)}
          placeholder="https://example.com"
        />
      </label>
      <button className="primary full" disabled={busy} onClick={create}>
        {busy ? 'Creating…' : 'Create Trackable QR'} <ScanLine size={17} />
      </button>
      {errorMsg && <div className="status error">{errorMsg}</div>}
      {created && (
        <div className="analytics-box">
          <div className="track-link">
            Tracking link:{' '}
            <code>
              {window.location.origin}/scan/{created.id}
            </code>
          </div>
          <div className="metric">
            <b>{scans.length}</b>
            <span>real scan events recorded across devices</span>
          </div>
          {scans.length === 0 && <p>No scans yet.</p>}
          <button
            className="secondary"
            onClick={async () => {
              const data = await QRCode.toDataURL(
                `${window.location.origin}/scan/${created.id}`,
                { width: 640, margin: 2 }
              );
              saveDataUrl(data, 'quicktools-trackable-qr.png');
            }}
          >
            Download Trackable QR
          </button>
        </div>
      )}
    </div>
  );
}

type Line = { description: string; quantity: number; price: number };
function BusinessDoc({ mode }: { mode: 'invoice' | 'quotation' }) {
  const [business, setBusiness] = useState('');
  const [customer, setCustomer] = useState('');
  const [lines, setLines] = useState<Line[]>([
    { description: '', quantity: 1, price: 0 },
  ]);
  const [tax, setTax] = useState(0);
  const [discount, setDiscount] = useState(0);
  const subtotal = lines.reduce((s, l) => s + Math.max(0, l.quantity) * Math.max(0, l.price), 0);
  const total = Math.max(
    0,
    subtotal + (subtotal * tax) / 100 - (subtotal * discount) / 100
  );
  function updateLine(i: number, k: keyof Line, v: string | number) {
    setLines(ls =>
      ls.map((l, n) =>
        n === i ? { ...l, [k]: typeof v === 'number' ? v : Number(v) } : l
      )
    );
  }
  function exportPdf() {
    const d = new jsPDF();
    d.setFontSize(22);
    d.text(mode === 'invoice' ? 'INVOICE' : 'QUOTATION', 18, 22);
    d.setFontSize(11);
    d.text(business || 'Business name', 18, 34);
    d.text(`Customer: ${customer || 'Customer'}`, 18, 42);
    let y = 58;
    lines.forEach((l, i) => {
      d.text(
        `${i + 1}. ${l.description || 'Item'} — ${l.quantity} × ${l.price.toFixed(2)} = ${(l.quantity * l.price).toFixed(2)}`,
        18,
        y
      );
      y += 8;
    });
    d.text(`Subtotal: ${subtotal.toFixed(2)}`, 18, y + 8);
    d.text(`Tax: ${tax}%`, 18, y + 16);
    d.text(`Discount: ${discount}%`, 18, y + 24);
    d.setFontSize(14);
    d.text(`Total: ${total.toFixed(2)}`, 18, y + 36);
    d.save(`quicktools-${mode}.pdf`);
  }
  return (
    <div className="business-grid">
      <div className="form-panel">
        <label>
          Business name
          <input value={business} onChange={e => setBusiness(e.target.value)} />
        </label>
        <label>
          Customer
          <input value={customer} onChange={e => setCustomer(e.target.value)} />
        </label>
        <div className="line-items">
          {lines.map((l, i) => (
            <div className="line-item" key={i}>
              <input
                placeholder="Item"
                value={l.description}
                onChange={e => updateLine(i, 'description', e.target.value)}
              />
              <input
                type="number"
                min="0"
                placeholder="Qty"
                value={l.quantity}
                onChange={e => updateLine(i, 'quantity', e.target.value)}
              />
              <input
                type="number"
                min="0"
                placeholder="Price"
                value={l.price}
                onChange={e => updateLine(i, 'price', e.target.value)}              />
            </div>
          ))}
        </div>
        <button
          className="secondary"
          onClick={() =>
            setLines([...lines, { description: '', quantity: 1, price: 0 }])
          }
        >
          + Add item
        </button>
        <div className="two-col">
          <label>
            Tax %
            <input
              type="number"
              min="0"
              value={tax}
              onChange={e => setTax(Number(e.target.value) || 0)}
            />
          </label>
          <label>
            Discount %
            <input
              type="number"
              min="0"
              value={discount}
              onChange={e => setDiscount(Number(e.target.value) || 0)}
            />
          </label>
        </div>
        <button className="primary full" onClick={exportPdf}>
          Download PDF <Download size={17} />
        </button>
      </div>
      <div className="business-preview">
        <span>{mode === 'invoice' ? 'INVOICE' : 'QUOTATION'}</span>
        <h2>{business || 'Your Business'}</h2>
        <p>Customer: {customer || 'Customer'}</p>
        {lines.map((l, i) => (
          <div key={i}>
            <span>{l.description || 'Item'}</span>
            <b>{(l.quantity * l.price).toFixed(2)}</b>
          </div>
        ))}
        <hr />
        <div>
          <span>Subtotal</span>
          <b>{subtotal.toFixed(2)}</b>
        </div>
        <div>
          <span>Tax</span>
          <b>{tax}%</b>
        </div>
        <div>
          <span>Discount</span>
          <b>{discount}%</b>
        </div>
        <strong className="grand-total">{total.toFixed(2)}</strong>
      </div>
    </div>
  );
}

function StaticPage({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) {
  return (
    <PageShell
      title={title}
      intro="QuickTools keeps this information clear, concise and easy to find."
      seoKey="home"
    >
      {children}
    </PageShell>
  );
}

function App() {
  const [route, setRoute] = useState(getRoute());
  useEffect(() => {
    const fn = () => setRoute(getRoute());
    window.addEventListener('popstate', fn);
    return () => window.removeEventListener('popstate', fn);
  }, []);
  useEffect(() => {
    const key = route.startsWith('tools/') ? route.split('/')[1] : route;
    const data = seo[key] || seo.home;
    document.title = data.title;
    let desc = document.querySelector('meta[name="description"]');
    if (!desc) {
      desc = document.createElement('meta');
      desc.setAttribute('name', 'description');
      document.head.appendChild(desc);
    }
    desc.setAttribute('content', data.description);
    let kw = document.querySelector('meta[name="keywords"]');
    if (!kw) {
      kw = document.createElement('meta');
      kw.setAttribute('name', 'keywords');
      document.head.appendChild(kw);
    }
    kw.setAttribute('content', data.keywords);

    // Give every public route its own canonical URL.
    const canonicalUrl = new URL(window.location.pathname, window.location.origin).href;
    let canonical = document.querySelector('link[rel="canonical"]');
    if (!canonical) {
      canonical = document.createElement('link');
      canonical.setAttribute('rel', 'canonical');
      document.head.appendChild(canonical);
    }
    canonical.setAttribute('href', canonicalUrl);

    let ogUrl = document.querySelector('meta[property="og:url"]');
    if (!ogUrl) {
      ogUrl = document.createElement('meta');
      ogUrl.setAttribute('property', 'og:url');
      document.head.appendChild(ogUrl);
    }
    ogUrl.setAttribute('content', canonicalUrl);
  }, [route]);
  const onNavigate = (r: string) => navigate(r);
  let content: React.ReactNode;
  if (route === 'home') content = <Home onNavigate={onNavigate} />;
  else if (route === 'tools')
    content = <ToolsDirectory onNavigate={onNavigate} />;
  else if (route === 'privacy')
    content = (
      <StaticPage title="Privacy Policy">
        <div className="content-card">
          <h2>Privacy at QuickTools</h2>
          <p>
            QuickTools aims to minimize unnecessary data collection. Many file
            tools process files directly in your browser and create results
            locally on your device.
          </p>
          <p>
            QuickTools may use cookies, local storage, analytics, advertising
            services, or other third-party technologies when those features are
            enabled. These technologies may process information such as device,
            browser, usage, or advertising-related data according to the
            applicable provider's policies.
          </p>
          <p>
            Do not upload confidential or sensitive information unless you are
            comfortable using the specific tool. If a feature changes how data
            is processed or stored, this policy should be updated to describe
            that behavior.
          </p>
        </div>
      </StaticPage>
    );
  else if (route === 'terms')
    content = (
      <StaticPage title="Terms of Service">
        <div className="content-card">
          <h2>Using QuickTools</h2>
          <p>
            By using QuickTools, you agree to use the service lawfully and
            responsibly and not to misuse, disrupt, or attempt to compromise
            the service.
          </p>
          <p>
            You are responsible for the files, links, text, and business
            information you provide. Tool outputs should be reviewed before
            being used for professional, financial, legal, or other important
            decisions.
          </p>
          <p>
            QuickTools is provided on an as-available basis. Features may be
            changed, improved, or discontinued as the service develops.
          </p>
        </div>
      </StaticPage>
    );
  else if (route === 'contact')
    content = (
      <StaticPage title="Contact QuickTools">
        <div className="content-card">
          <h2>Feedback & Support</h2>
          <p>
            We welcome feedback about broken tools, bugs, usability issues, and
            feature requests.
          </p>
          <p>
            A dedicated support email or contact form should be added here
            before advertising a public support channel. Until then, users can
            use the available platform contact method shown on the website.
          </p>
        </div>
      </StaticPage>
    );
  else if (route.startsWith('scan/')) {
    content = <QRRedirectPage id={route.split('/')[1]} />;
  }
  else if (route.startsWith('tools/'))
    content = (
      <ToolPage
        toolKey={route.split('/')[1] as ToolKey}
        onNavigate={onNavigate}
      />
    );
  else content = <Home onNavigate={onNavigate} />;
  return (
    <><StructuredData route={route} /><Layout route={route} onNavigate={onNavigate}>
      {content}
    </Layout></>
  );
}

function formatBytes(n: number) {
  if (!n) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(n) / Math.log(1024));
  return `${(n / Math.pow(1024, i)).toFixed(i ? 1 : 0)} ${units[i]}`;
}
function saveBlob(blob: Blob, name: string) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
function saveDataUrl(data: string, name: string) {
  const a = document.createElement('a');
  a.href = data;
  a.download = name;
  a.click();
}
function QRRedirectPage({ id }: { id: string }) {
  const [message, setMessage] = useState('Opening destination…');
  useEffect(() => {
    let cancelled = false;
    const redirect = async () => {
      try {
        const response = await api.get(`/api/qr/${id}/redirect`);
        const target = response.data?.target;
        if (!target) throw new Error('NOT_FOUND');
        if (!cancelled) window.location.replace(target);
      } catch {
        if (!cancelled) setMessage('This QR tracking link is invalid or no longer available.');
      }
    };
    void redirect();
    return () => {
      cancelled = true;
    };
  }, [id]);
  return <StaticPage title="Redirecting…"><div className="content-card"><p>{message}</p></div></StaticPage>;
}

async function downloadSvg(value: string) {
  const svg = await QRCode.toString(value, {
    type: 'svg',
    width: 640,
    margin: 2,
  });
  const blob = new Blob([svg], { type: 'image/svg+xml' });
  saveBlob(blob, 'quicktools-qr.svg');
}

export default App;