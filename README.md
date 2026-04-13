# PoterPointAI

AI-Powered Presentation Builder — [ppt.poterai.com](https://ppt.poterai.com)

## Features
- AI slide generation (Google Gemini)
- 8 slide types with drag-and-drop canvas editor
- Template library (6 built-in + custom)
- Light/Dark mode with 5 color themes
- PPTX, PDF, JSON export
- Multi-selection, grouping, inline editing
- Photo management, Excel sync, data masking

## Quick Start

```bash
npm install
cp .env.example .env  # Add your GOOGLE_API_KEY
npm run dev            # Frontend: http://localhost:5173
npm run server         # Backend: http://localhost:3001
```

## Production

```bash
npm run build          # Build frontend
npm start              # Start production server
```

## Environment Variables

```
GOOGLE_API_KEY=your-gemini-api-key
PORT=3001
```

## Tech Stack
React 19 | Vite 8 | Express 5 | Google Gemini AI | PptxGenJS | jsPDF
