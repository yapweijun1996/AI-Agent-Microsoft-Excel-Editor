# AI Excel Editor - Multi-Agent Spreadsheet Automation

![AI Excel Editor](https://img.shields.io/badge/AI-Excel%20Editor-blue?style=for-the-badge&logo=microsoftexcel)
![OpenAI](https://img.shields.io/badge/OpenAI-GPT--4-green?style=flat&logo=openai)
![Gemini](https://img.shields.io/badge/Google-Gemini-orange?style=flat&logo=google)
![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-yellow?style=flat&logo=javascript)

A powerful web-based Excel editor with multi-agent AI automation that runs entirely in your browser. No server required!

## ✨ Features

### 🤖 Multi-Agent AI System
- **Planner Agent**: Breaks down complex requests into manageable tasks
- **Executor Agent**: Performs precise spreadsheet operations
- **Validator Agent**: Ensures data integrity and safety
- **Multiple AI Providers**: OpenAI GPT-4 and Google Gemini support

### 📊 Advanced Spreadsheet Features
- **Full Excel Compatibility**: Import/export .xlsx and .csv files
- **Multi-Sheet Support**: Create, manage, and switch between multiple sheets
- **Real-time Editing**: Live cell editing with instant updates
- **Undo/Redo**: Complete history management with 50-level undo
- **Keyboard Shortcuts**: Professional shortcuts for efficient editing

### 🎯 Smart Automation
- **Natural Language Commands**: "Add totals row", "Format as currency", etc.
- **Task Management**: Visual task tracking with status updates
- **Dry Run Mode**: Preview AI changes before applying
- **Context-Aware**: AI understands your current sheet structure

### 🔒 Privacy & Security
- **Client-Side Only**: All processing happens in your browser
- **Your API Keys**: Use your own OpenAI/Gemini API keys
- **Local Storage**: Data persists locally using IndexedDB
- **No Server**: Zero data transmission to external servers

## 🚀 Quick Start

### 1. Open the Application
```bash
# Clone the repository
git clone https://github.com/your-repo/AI-Agent-Microsoft-Excel-Editor
cd AI-Agent-Microsoft-Excel-Editor

# Open in your browser
open web/index.html
```

### 2. Set Your API Keys
- Click "Set OpenAI Key" or "Set Gemini Key" in the header
- Enter your API key (get one from [OpenAI](https://platform.openai.com/) or [Google AI](https://ai.google.dev/))
- Choose whether to persist the key locally

### 3. Start Using AI
Try these example commands:
- "Create a header row with Name, Age, Email, Salary"
- "Add a totals row that sums the salary column"
- "Format column D as currency"
- "Sort the data by age in ascending order"
- "Add a new sheet for expenses"

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+S` | Export as XLSX |
| `Ctrl+O` | Import XLSX file |
| `Ctrl+Z` | Undo |
| `Ctrl+Y` | Redo |
| `Ctrl+T` | Add new sheet |
| `Ctrl+W` | Delete current sheet |
| `F2` | Focus chat input |
| `Tab` | Switch between sheets |
| `Ctrl+1-9` | Switch to sheet by number |

## 🏗️ Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Planner       │    │   Executor      │    │   Validator     │
│   Agent         │────▶│   Agent         │────▶│   Agent         │
└─────────────────┘    └─────────────────┘    └─────────────────┘
         │                       │                       │
         ▼                       ▼                       ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Frontend Application                         │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐  │
│  │ Spreadsheet │  │    Chat     │  │     Task Management     │  │
│  │   Editor    │  │   Interface │  │        System           │  │
│  └─────────────┘  └─────────────┘  └─────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
         │                       │                       │
         ▼                       ▼                       ▼
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   IndexedDB     │    │   localStorage  │    │   SheetJS       │
│   Storage       │    │   Settings      │    │   Engine        │
└─────────────────┘    └─────────────────┘    └─────────────────┘
```

## 🛠️ Technology Stack

- **Frontend**: Vanilla JavaScript (ES6+), HTML5, CSS3
- **Styling**: Tailwind CSS
- **Spreadsheet Engine**: SheetJS (XLSX)
- **Storage**: IndexedDB + localStorage
- **AI Providers**: OpenAI GPT-4, Google Gemini
- **Build**: No build process required - pure web standards

## 📱 Browser Compatibility

- ✅ Chrome 60+
- ✅ Firefox 55+
- ✅ Safari 11+
- ✅ Edge 79+

## 🔧 Development

### File Structure
```
web/
├── index.html          # Main HTML file
├── styles.css          # Enhanced CSS styles
├── app.js             # Core application logic
└── sample-data.xlsx   # Optional example file
```

### Core Components
- **Modal System**: Reusable modal dialogs
- **Toast Notifications**: User feedback system
- **Task Manager**: AI task coordination
- **Sheet Renderer**: Dynamic spreadsheet display
- **History Manager**: Undo/redo functionality

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙋 Support

- 📚 **Documentation**: See [DEVELOPER_BOOK.md](DEVELOPER_BOOK.md) for detailed implementation guide
- 🐛 **Issues**: Report bugs or request features via GitHub Issues
- 💡 **Discussions**: Join community discussions for ideas and help

## ⚠️ Security Notice

This application runs AI calls directly from the browser using your API keys. For production use, consider implementing a server-side proxy to protect API keys. The current architecture is ideal for:
- Personal use and prototyping
- Internal tools where users provide their own keys
- Educational and demonstration purposes

---

**Made with ❤️ and AI** - Democratizing spreadsheet automation with artificial intelligence.