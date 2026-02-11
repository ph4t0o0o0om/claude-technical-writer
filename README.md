# Claude Technical Writer

An AI-powered multi-agent system that automatically generates comprehensive user manuals for web applications. It uses the [Claude Agent SDK](https://github.com/anthropics/claude-agent-sdk) to orchestrate a pipeline of specialized agents that research, document, and verify technical documentation — complete with screenshots and diagrams.

The default target application is [OWASP Juice Shop](https://owasp.org/www-project-juice-shop/), a deliberately insecure e-commerce web app used for security training, but the approach can be adapted to other web applications.

## How It Works

The system orchestrates three agents in sequence:

1. **Research Agent** — Navigates the target web application like a real user using browser automation. It explores features, captures screenshots at key points, and gathers information about the platform's capabilities and workflows.

2. **Technical Writer Agent** — Takes the research agent's findings and writes a structured, modular user manual in Markdown. Screenshots and Mermaid-based diagrams (exported as PNG) are embedded inline for a textbook-like reading experience.

3. **Verifier Agent** — Reviews the final documentation for accuracy, verifying that screenshots match their corresponding sections, fixing broken image links, and removing any inaccurate or exaggerated content.

The output is a set of Markdown files and images written to the `output/` directory.

## Prerequisites

- **Python 3.13+**
- **[uv](https://docs.astral.sh/uv/)** — Fast Python package manager
- **Node.js / npm** — Required for the MCP memory server (`@modelcontextprotocol/server-memory`)
- **OWASP Juice Shop** running locally on `http://localhost:3001` (see [Running Juice Shop](#running-juice-shop) below)
- **API credentials** for an LLM provider (Anthropic-compatible API)

## Setup

### 1. Clone the repository

```bash
git clone <repo-url>
cd claude-technical-writer
```

### 2. Install Python dependencies

```bash
uv sync
```

This installs the project dependencies defined in `pyproject.toml`, including:
- `claude-agent-sdk` — SDK for orchestrating Claude-based agents
- `loguru` — Structured logging

### 3. Configure environment variables

Create a `.env` file in the project root with the following variables:

```
ANTHROPIC_MODEL="<model-name>"
ANTHROPIC_AUTH_TOKEN="<your-api-key>"
ANTHROPIC_API_KEY="<your-anthropic-api-key>"
ANTHROPIC_DEFAULT_SONNET_MODEL="<model-name>"
ANTHROPIC_DEFAULT_OPUS_MODEL="<model-name>"
ANTHROPIC_DEFAULT_HAIKU_MODEL="<model-name>"
ANTHROPIC_BASE_URL="<api-base-url>"
```

If using Anthropic directly, set `ANTHROPIC_API_KEY` and leave `ANTHROPIC_BASE_URL` as the default. If using a proxy like OpenRouter, configure `ANTHROPIC_AUTH_TOKEN` and `ANTHROPIC_BASE_URL` accordingly.

### 4. Prepare the input file

Edit `INPUT.md` to describe the target application. This file should contain:
- A high-level description of the application
- The URL where the application is running
- Any credentials or signup instructions needed to access the app

The default `INPUT.md` is pre-configured for OWASP Juice Shop running at `http://localhost:3001`.

### Running Juice Shop

The easiest way to run OWASP Juice Shop locally is with Docker:

```bash
docker run -d -p 3001:3000 bkimminich/juice-shop
```

Verify it's running by visiting `http://localhost:3001` in your browser.

## Usage

Run the agent:

```bash
uv run python main.py
```

The system will:
1. Read `INPUT.md` for application details and access instructions
2. Launch the research agent to browse the application and capture screenshots
3. Hand off findings to the technical writer agent to produce the user manual
4. Run the verifier agent to review and correct the documentation
5. Write the final output (Markdown files, screenshots, and diagrams) to the `output/` directory

> **Note:** This process can take a significant amount of time depending on the complexity of the target application, as the agents perform real browser interactions and multi-turn LLM conversations. The `max_turns` is set to 10,000 to allow thorough exploration.

## Output

All generated documentation is written to the `output/` directory:

```
output/
├── 00_index.md              # Table of contents
├── 01_getting_started.md    # Getting started guide
├── 02_account_management.md # Account and profile management
├── 03_product_browsing.md   # Browsing and searching products
├── ...                      # Additional feature guides
├── *.png                    # Screenshots captured during research
└── diagrams/                # Mermaid diagrams exported as PNG
```

The `output/` directory contents are git-ignored (only `.gitkeep` is tracked) since they are generated artifacts.

## Project Structure

```
claude-technical-writer/
├── main.py           # Entry point — agent definitions and orchestration
├── INPUT.md          # Target application description and access details
├── pyproject.toml    # Python project metadata and dependencies
├── .python-version   # Python version (3.13)
├── .env              # Environment variables (not committed)
├── .gitignore        # Git ignore rules
├── output/           # Generated documentation output (git-ignored)
│   └── .gitkeep
└── README.md         # This file
```

## Customization

To generate documentation for a different web application:

1. Update `INPUT.md` with the target application's description, URL, and credentials
2. Modify the agent prompts in `main.py` to reference the new application instead of OWASP Juice Shop
3. Adjust the `APPROVED_TOOLS` list if your agents need additional capabilities
4. Run `uv run python main.py`

## License

See the project repository for license information.
