import React, { useEffect, useState, useRef } from 'react';
import { useNavigate } from 'react-router-dom';
import { marked } from 'marked';
import { ArrowLeft, BookOpen, Loader } from 'lucide-react';
import { apiUrl } from '../utils/auth';

// Configure marked to open external links in a new tab and to be GFM-compatible
marked.setOptions({
    gfm: true,
    breaks: true,
});

// Custom renderer to make anchor links work within the page
const renderer = new marked.Renderer();
renderer.heading = ({ text, depth }) => {
    const slug = text
        .toLowerCase()
        .replace(/[^\w\s-]/g, '')
        .trim()
        .replace(/\s+/g, '-');
    return `<h${depth} id="${slug}">${text}</h${depth}>\n`;
};
renderer.link = ({ href, text }) => {
    // Internal anchor links stay in-page; others open in a new tab
    if (href && href.startsWith('#')) {
        return `<a href="${href}">${text}</a>`;
    }
    return `<a href="${href}" target="_blank" rel="noopener noreferrer">${text}</a>`;
};

export default function Help() {
    const navigate = useNavigate();
    const [markdown, setMarkdown] = useState('');
    const [html, setHtml] = useState('');
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState('');
    const contentRef = useRef(null);

    useEffect(() => {
        setLoading(true);
        fetch(apiUrl('/api/public/readme'))
            .then(res => {
                if (!res.ok) throw new Error(`Failed to load help content (${res.status})`);
                return res.text();
            })
            .then(text => {
                setMarkdown(text);
                const parsed = marked.parse(text, { renderer });
                setHtml(parsed);
                setLoading(false);
            })
            .catch(err => {
                setError(err.message);
                setLoading(false);
            });
    }, []);

    // Smooth-scroll to anchor when clicking in-page TOC links
    useEffect(() => {
        if (!contentRef.current) return;
        const handleClick = (e) => {
            const a = e.target.closest('a[href^="#"]');
            if (!a) return;
            e.preventDefault();
            const id = decodeURIComponent(a.getAttribute('href').slice(1));
            const target = document.getElementById(id);
            if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        };
        const el = contentRef.current;
        el.addEventListener('click', handleClick);
        return () => el.removeEventListener('click', handleClick);
    }, [html]);

    return (
        <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
            {/* Navbar */}
            <nav className="navbar navbar-expand-lg navbar-light" style={{ backgroundColor: 'skyblue' }}>
                <div className="container-fluid">
                    <span className="navbar-brand d-flex align-items-center">
                        <BookOpen className="me-2" size={20} /> Help &amp; Documentation
                    </span>
                    <div className="d-flex align-items-center">
                        <button
                            onClick={() => navigate(-1)}
                            className="btn btn-outline-dark btn-sm d-flex align-items-center"
                            title="Go Back"
                        >
                            <ArrowLeft size={14} className="me-1" /> Back
                        </button>
                    </div>
                </div>
            </nav>

            {/* Content */}
            <main className="flex-1 w-full mx-auto px-4 sm:px-6 py-8" style={{ maxWidth: 900 }}>
                {loading && (
                    <div className="d-flex align-items-center justify-content-center py-5 text-muted">
                        <Loader size={24} className="me-2 animate-spin" />
                        Loading documentation…
                    </div>
                )}

                {error && (
                    <div className="alert alert-danger" role="alert">
                        <strong>Error:</strong> {error}
                    </div>
                )}

                {!loading && !error && (
                    <div
                        ref={contentRef}
                        className="bg-white border border-gray-200 rounded-2xl shadow-sm p-6 p-md-8"
                        style={{ lineHeight: 1.75 }}
                        dangerouslySetInnerHTML={{ __html: html }}
                    />
                )}
            </main>

            {/* Inline styles for markdown rendering */}
            <style>{`
                /* Headings */
                .bg-white h1 { font-size: 2rem; font-weight: 700; margin-top: 0; margin-bottom: 0.75rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 0.4rem; }
                .bg-white h2 { font-size: 1.5rem; font-weight: 600; margin-top: 2rem; margin-bottom: 0.5rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.3rem; }
                .bg-white h3 { font-size: 1.2rem; font-weight: 600; margin-top: 1.5rem; margin-bottom: 0.4rem; }
                .bg-white h4, .bg-white h5, .bg-white h6 { font-size: 1rem; font-weight: 600; margin-top: 1.2rem; margin-bottom: 0.3rem; }

                /* Paragraphs & lists */
                .bg-white p { margin-bottom: 0.75rem; }
                .bg-white ul, .bg-white ol { padding-left: 1.5rem; margin-bottom: 0.75rem; }
                .bg-white li { margin-bottom: 0.25rem; }

                /* Code */
                .bg-white code { background: #f1f5f9; border-radius: 4px; padding: 0.15em 0.4em; font-size: 0.875rem; color: #be123c; }
                .bg-white pre { background: #1e293b; color: #e2e8f0; border-radius: 8px; padding: 1rem; overflow-x: auto; margin-bottom: 1rem; }
                .bg-white pre code { background: transparent; color: inherit; padding: 0; font-size: 0.85rem; }

                /* Tables */
                .bg-white table { width: 100%; border-collapse: collapse; margin-bottom: 1rem; font-size: 0.9rem; }
                .bg-white th { background: #f8fafc; font-weight: 600; text-align: left; padding: 0.5rem 0.75rem; border: 1px solid #e2e8f0; }
                .bg-white td { padding: 0.45rem 0.75rem; border: 1px solid #e2e8f0; vertical-align: top; }
                .bg-white tr:nth-child(even) td { background: #f8fafc; }

                /* Block quotes */
                .bg-white blockquote { border-left: 4px solid #94a3b8; background: #f8fafc; margin: 0.5rem 0 1rem 0; padding: 0.5rem 1rem; border-radius: 0 6px 6px 0; color: #475569; }

                /* Links */
                .bg-white a { color: #2563eb; text-decoration: none; }
                .bg-white a:hover { text-decoration: underline; }

                /* Horizontal rule */
                .bg-white hr { border: none; border-top: 1px solid #e2e8f0; margin: 1.5rem 0; }

                /* Animated spinner */
                @keyframes spin { to { transform: rotate(360deg); } }
                .animate-spin { animation: spin 1s linear infinite; }
            `}</style>
        </div>
    );
}
