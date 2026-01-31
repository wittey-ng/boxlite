# DocGenerator Image - Document Generation Container
#
# Provides tools for generating business documents:
# - PPT (python-pptx)
# - Word (python-docx)
# - Excel (openpyxl, pandas)
# - PDF (reportlab)
# - Charts (matplotlib, plotly)
# - Vector DB clients (chromadb, pinecone, weaviate)
# - AI APIs (openai, anthropic)

FROM python:3.11-slim

# Install system dependencies for document generation
RUN apt-get update && apt-get install -y --no-install-recommends \
    libcairo2-dev \
    libpango1.0-dev \
    libjpeg-dev \
    libgif-dev \
    librsvg2-dev \
    wkhtmltopdf \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# Install Python packages for document generation
RUN pip install --no-cache-dir \
    # Document generation
    python-pptx \
    python-docx \
    openpyxl \
    pandas \
    reportlab \
    pdfkit \
    # Data visualization
    matplotlib \
    plotly \
    seaborn \
    pillow \
    # Vector database clients
    chromadb \
    # AI API clients
    openai \
    anthropic \
    # Utilities
    python-dotenv \
    numpy

WORKDIR /workspace
