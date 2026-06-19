# Excel Duty Card Generator

Police Duty Card Generator for UP Police Bulandshahr — generates print-ready A5 landscape duty cards from Excel (.xlsx) duty chart files.

## Key Features

- **Excel-based extraction**: Upload `.xlsx` duty charts directly — no AI/API calls needed
- **Client-side parsing**: All data extraction happens in the browser using SheetJS
- **Zero API dependencies**: No Gemini, OpenRouter, or any external API keys required
- **Print-ready cards**: Generates pixel-perfect A5 landscape duty cards
- **Edit & Search**: Edit any card's details, search across all cards
- **Hindi/Devanagari support**: Full support for Hindi text and police rank abbreviations

## Tech Stack

- **Next.js 16** with App Router
- **SheetJS (xlsx)** for client-side Excel parsing
- **Tailwind CSS 4** for styling
- **TypeScript** for type safety

## Getting Started

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

## How It Works

1. Upload your Excel (.xlsx) duty chart file
2. The app parses the file client-side using SheetJS
3. Officer records are extracted and mapped to duty cards
4. Edit, search, and print individual or all cards

## No API Keys Required

Unlike the previous version, this project has **zero external API dependencies**. Everything runs locally in the browser.
