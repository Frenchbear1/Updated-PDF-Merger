# Updated File Merger

A desktop Electron app that merges PDFs, PowerPoints, and images.

## Features

- PDF-only batches use the original qpdf-based merge engine (fast mode + safe mode for large inputs).
- PowerPoint/image batches use the PPTX merge pipeline and can export to `.pptx` or `.pdf`.
- Mixed batches (PDF + PowerPoint/image) are converted and merged into a single `.pdf`.
- Drag-and-drop file picking and drag-to-reorder list control.
- Image preview in the file list.
- Progress updates during long merges.
- Includes bundled qpdf setup/usage for robust PDF operations.

## Notes

- Mixed batches (PDF + PowerPoint/image) are supported when output is PDF.
