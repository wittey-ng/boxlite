"""
DocGeneratorBox - Document Generation Container.

Provides tools for generating business analysis documents from data:
- PowerPoint presentations
- Word reports
- Excel spreadsheets
- PDF documents
- Data visualizations
- Vector database integration
"""

from dataclasses import dataclass
from typing import Optional, List, Dict, Any, Literal, TYPE_CHECKING
import json
import base64

from .simplebox import SimpleBox
from .errors import ExecError

if TYPE_CHECKING:
    from .boxlite import Boxlite

__all__ = ["DocGeneratorBox", "DocGeneratorBoxOptions"]


@dataclass
class DocGeneratorBoxOptions:
    """
    Configuration for DocGeneratorBox.

    Example:
        >>> opts = DocGeneratorBoxOptions(
        ...     cpu=2,
        ...     memory=4096,
        ...     vector_db_type="chromadb",
        ...     vector_db_host="localhost"
        ... )
        >>> async with DocGeneratorBox(opts) as docgen:
        ...     await docgen.generate_ppt("Report", [{"title": "..."}], "/output/report.pptx")
    """
    cpu: int = 2
    memory: int = 4096  # Memory in MiB
    image: str = "boxlite/doc-generator:latest"
    vector_db_type: Optional[Literal["chromadb"]] = None  # Only ChromaDB currently supported
    vector_db_host: Optional[str] = None
    vector_db_port: Optional[int] = None


class DocGeneratorBox(SimpleBox):
    """
    Specialized container for document generation.

    Provides methods to generate various document formats from data,
    with optional vector database and AI integration.

    Usage:
        >>> async with DocGeneratorBox() as docgen:
        ...     # Generate PowerPoint
        ...     await docgen.generate_ppt("Report", slides, "/output/report.pptx")
        ...
        ...     # Generate Excel
        ...     await docgen.generate_excel(sheets_dict, "/output/data.xlsx")
    """

    # Valid chart types
    _CHART_TYPES = {"bar", "line", "pie", "scatter"}

    # Valid templates
    _PPT_TEMPLATES = {"business", "minimal", "corporate"}

    def __init__(
            self,
            options: Optional[DocGeneratorBoxOptions] = None,
            runtime: Optional['Boxlite'] = None,
            **kwargs
    ):
        """
        Create a DocGeneratorBox.

        Args:
            options: Document generator configuration (uses defaults if None)
            runtime: Optional runtime instance (uses global default if None)
            **kwargs: Additional configuration options (volumes, env, etc.)
        """
        opts = options or DocGeneratorBoxOptions()

        self._vector_db_type = opts.vector_db_type
        self._vector_db_host = opts.vector_db_host
        self._vector_db_port = opts.vector_db_port

        super().__init__(
            image=opts.image,
            memory_mib=opts.memory,
            cpus=opts.cpu,
            runtime=runtime,
            **kwargs
        )

    async def _exec_python_with_data(self, script: str, data: Dict[str, Any]) -> str:
        """
        Execute Python script with JSON data passed via environment variable.

        Args:
            script: Python script to execute
            data: Data dictionary to pass to script

        Returns:
            Script stdout output

        Raises:
            ExecError: If script exits with non-zero code
        """
        # Encode data as base64 JSON to safely pass through environment
        json_data = json.dumps(data)
        b64_data = base64.b64encode(json_data.encode()).decode()

        # Prepend data decoding to script
        full_script = f"""
import json
import base64
import os

# Decode data from environment
_data_b64 = os.environ.get('DOCGEN_DATA', '')
_data_json = base64.b64decode(_data_b64).decode()
data = json.loads(_data_json)

# User script
{script}
"""

        # Execute with data in environment
        result = await self.exec(
            "python", "-c", full_script,
            env={"DOCGEN_DATA": b64_data}
        )

        # Check for errors
        if result.exit_code != 0:
            raise ExecError(
                f"Script failed with exit code {result.exit_code}\\n"
                f"Stdout: {result.stdout}\\n"
                f"Stderr: {result.stderr}"
            )

        return result.stdout.strip()

    @staticmethod
    def _validate_safe_path(path: str, allow_absolute: bool = False) -> None:
        """
        Validate path is safe (no directory traversal attacks).

        Args:
            path: Path to validate
            allow_absolute: Whether to allow absolute paths

        Raises:
            ValueError: If path contains unsafe components
        """
        import os

        # Reject paths with .. (parent directory traversal)
        if '..' in path.split(os.sep):
            raise ValueError(f"Path traversal detected: {path}")

        # Reject absolute paths if not allowed
        if not allow_absolute and os.path.isabs(path):
            raise ValueError(f"Absolute paths not allowed: {path}")

    @staticmethod
    def _validate_workspace_name(name: str) -> None:
        """
        Validate workspace name is a single path segment.

        Args:
            name: Workspace name to validate

        Raises:
            ValueError: If name contains path separators or unsafe components
        """
        import os

        # Reject absolute paths
        if os.path.isabs(name):
            raise ValueError(f"Workspace name cannot be absolute path: {name}")

        # Reject names with path separators
        if os.sep in name or (os.altsep and os.altsep in name):
            raise ValueError(f"Workspace name cannot contain path separators: {name}")

        # Reject names with ..
        if '..' in name:
            raise ValueError(f"Workspace name cannot contain '..': {name}")

        # Reject empty names
        if not name or name.strip() == '':
            raise ValueError("Workspace name cannot be empty")

    # ==================== PowerPoint Generation ====================

    async def generate_ppt(
            self,
            title: str,
            slides_data: List[Dict[str, Any]],
            output_path: str,
            template: str = "business"
    ) -> str:
        """
        Generate PowerPoint presentation.

        Args:
            title: Presentation title
            slides_data: List of slide configurations, each with 'title' and 'content'
            output_path: Output file path inside container
            template: Template style ("business", "minimal", "corporate")

        Returns:
            Success message with output path

        Raises:
            ValueError: If template is invalid
            ExecError: If generation fails

        Example:
            >>> slides = [
            ...     {"title": "Overview", "content": {"text": "Summary..."}},
            ...     {"title": "Data", "content": {"chart": {...}}}
            ... ]
            >>> await docgen.generate_ppt("Report", slides, "/output/report.pptx")
        """
        if template not in self._PPT_TEMPLATES:
            raise ValueError(f"Invalid template: {template}. Must be one of {self._PPT_TEMPLATES}")

        script = """
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Get parameters from data
title = data['title']
slides_data = data['slides']
template = data['template']
output_path = data['output_path']

# Template colors
if template == "business":
    title_color = RGBColor(31, 78, 121)
    accent_color = RGBColor(68, 114, 196)
elif template == "minimal":
    title_color = RGBColor(50, 50, 50)
    accent_color = RGBColor(100, 100, 100)
else:
    title_color = RGBColor(0, 0, 0)
    accent_color = RGBColor(50, 50, 50)

# Title slide
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title_shape = slide.shapes.title
subtitle = slide.placeholders[1]

title_shape.text = title
title_shape.text_frame.paragraphs[0].font.color.rgb = title_color
title_shape.text_frame.paragraphs[0].font.size = Pt(44)

subtitle.text = "Generated by DocGeneratorBox"
subtitle.text_frame.paragraphs[0].font.color.rgb = accent_color

# Content slides
for slide_info in slides_data:
    blank_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = slide_info.get('title', '')
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = title_color

    content = slide_info.get('content', {})
    if 'text' in content:
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
        text_frame = text_box.text_frame
        text_frame.text = content['text']
        text_frame.paragraphs[0].font.size = Pt(18)

prs.save(output_path)
print(f"PPT generated: {output_path}")
"""

        data = {
            "title": title,
            "slides": slides_data,
            "template": template,
            "output_path": output_path
        }

        return await self._exec_python_with_data(script, data)

    # ==================== Word Generation ====================

    async def generate_word_report(
            self,
            title: str,
            sections: List[Dict[str, Any]],
            output_path: str
    ) -> str:
        """
        Generate Word document report.

        Args:
            title: Document title
            sections: List of sections with 'heading', 'paragraphs', 'table', etc.
            output_path: Output file path inside container

        Returns:
            Success message with output path

        Raises:
            ExecError: If generation fails

        Example:
            >>> sections = [
            ...     {"heading": "Summary", "paragraphs": ["Analysis..."]},
            ...     {"heading": "Data", "table": [["Col1", "Col2"], ["A", "B"]]}
            ... ]
            >>> await docgen.generate_word_report("Report", sections, "/output/report.docx")
        """
        script = """
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

style = doc.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(11)

# Get parameters
title = data['title']
sections = data['sections']
output_path = data['output_path']

title_heading = doc.add_heading(title, 0)
title_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

for section in sections:
    doc.add_heading(section.get('heading', ''), level=1)

    if 'paragraphs' in section:
        for para_text in section['paragraphs']:
            p = doc.add_paragraph(para_text)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if 'table' in section:
        table_data = section['table']
        if table_data and len(table_data) > 0:
            # Validate table dimensions
            max_cols = max(len(row) for row in table_data)
            table = doc.add_table(rows=len(table_data), cols=max_cols)
            table.style = 'Light Grid Accent 1'

            for i, row in enumerate(table_data):
                for j, cell_value in enumerate(row):
                    table.rows[i].cells[j].text = str(cell_value)

    if section.get('page_break', False):
        doc.add_page_break()

doc.save(output_path)
print(f"Word document generated: {output_path}")
"""

        data = {
            "title": title,
            "sections": sections,
            "output_path": output_path
        }

        return await self._exec_python_with_data(script, data)

    # ==================== Excel Generation ====================

    async def generate_excel_report(
            self,
            sheets: Dict[str, List[Dict[str, Any]]],
            output_path: str
    ) -> str:
        """
        Generate Excel spreadsheet.

        Args:
            sheets: Dictionary of sheet names to list of row dictionaries
            output_path: Output file path inside container

        Returns:
            Success message with output path

        Raises:
            ExecError: If generation fails
            ValueError: If sheet names are invalid (>31 chars or forbidden chars)

        Example:
            >>> sheets = {
            ...     "Products": [
            ...         {"name": "Product A", "price": 100},
            ...         {"name": "Product B", "price": 200}
            ...     ]
            ... }
            >>> await docgen.generate_excel_report(sheets, "/output/data.xlsx")
        """
        # Validate sheet names (Excel limits)
        for sheet_name in sheets.keys():
            if len(sheet_name) > 31:
                raise ValueError(f"Sheet name too long (max 31 chars): {sheet_name}")
            if any(c in sheet_name for c in [':', '\\\\', '/', '?', '*', '[', ']']):
                raise ValueError(f"Sheet name contains forbidden characters: {sheet_name}")

        script = """
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

sheets_data = data['sheets']
output_path = data['output_path']

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for sheet_name, sheet_data in sheets_data.items():
        df = pd.DataFrame(sheet_data)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

wb = load_workbook(output_path)

header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True, size=12)

for sheet in wb.worksheets:
    # Style header row
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Auto-adjust column widths
    for column_cells in sheet.columns:
        length = max(len(str(cell.value or '')) for cell in column_cells)
        sheet.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 50)

wb.save(output_path)
print(f"Excel generated: {output_path}")
"""

        data = {
            "sheets": sheets,
            "output_path": output_path
        }

        return await self._exec_python_with_data(script, data)

    # ==================== Chart Generation ====================

    async def generate_chart(
            self,
            data_dict: Dict[str, List],
            chart_type: Literal["bar", "line", "pie", "scatter"],
            output_path: str,
            title: str = ""
    ) -> str:
        """
        Generate chart image.

        Args:
            data_dict: Chart data (format depends on chart_type)
                - bar/pie: {"labels": [...], "values": [...]}
                - line/scatter: {"x": [...], "y": [...]}
            chart_type: Type of chart
            output_path: Output image path inside container
            title: Chart title

        Returns:
            Success message with output path

        Raises:
            ValueError: If chart_type is invalid
            ExecError: If generation fails

        Example:
            >>> data = {"labels": ["A", "B", "C"], "values": [10, 20, 30]}
            >>> await docgen.generate_chart(data, "bar", "/output/chart.png", "Sales")
        """
        if chart_type not in self._CHART_TYPES:
            raise ValueError(f"Invalid chart_type: {chart_type}. Must be one of {self._CHART_TYPES}")

        script = """
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# Configure CJK font support
plt.rcParams['font.sans-serif'] = ['Noto Sans CJK SC', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False  # Fix minus sign display

chart_data = data['chart_data']
chart_type = data['chart_type']
output_path = data['output_path']
title = data['title']

fig, ax = plt.subplots(figsize=(10, 6))

if chart_type == "bar":
    ax.bar(chart_data['labels'], chart_data['values'])
elif chart_type == "line":
    ax.plot(chart_data['x'], chart_data['y'], marker='o')
elif chart_type == "pie":
    ax.pie(chart_data['values'], labels=chart_data['labels'], autopct='%1.1f%%')
elif chart_type == "scatter":
    ax.scatter(chart_data['x'], chart_data['y'])

if title:
    ax.set_title(title, fontsize=16, fontweight='bold')

ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(output_path, dpi=300, bbox_inches='tight')
plt.close()

print(f"Chart saved: {output_path}")
"""

        data = {
            "chart_data": data_dict,
            "chart_type": chart_type,
            "output_path": output_path,
            "title": title
        }

        return await self._exec_python_with_data(script, data)

    # ==================== Vector Database Query ====================

    async def query_vector_db(
            self,
            collection: str,
            query: str,
            n_results: int = 10
    ) -> Dict[str, Any]:
        """
        Query vector database (ChromaDB only).

        Args:
            collection: Collection name
            query: Query text
            n_results: Number of results to return

        Returns:
            Query results as dictionary

        Raises:
            ValueError: If vector_db_type not configured or unsupported
            ExecError: If query fails

        Note:
            Requires vector_db_type="chromadb", vector_db_host, vector_db_port to be set
        """
        if not self._vector_db_type:
            raise ValueError("vector_db_type not configured in DocGeneratorBoxOptions")

        if self._vector_db_type != "chromadb":
            raise ValueError(f"Unsupported vector_db_type: {self._vector_db_type}. Only 'chromadb' is currently supported.")

        script = """
import chromadb

host = data['host']
port = data['port']
collection_name = data['collection']
query_text = data['query']
n_results = data['n_results']

client = chromadb.HttpClient(host=host, port=port)
collection = client.get_collection(collection_name)

results = collection.query(
    query_texts=[query_text],
    n_results=n_results
)

import json
print(json.dumps(results))
"""

        data = {
            "host": self._vector_db_host or "localhost",
            "port": self._vector_db_port or 8000,
            "collection": collection,
            "query": query,
            "n_results": n_results
        }

        result = await self._exec_python_with_data(script, data)
        return json.loads(result)

    # ==================== AI Content Generation ====================

    async def generate_summary(
            self,
            text: str,
            prompt_template: Optional[str] = None
    ) -> str:
        """
        Generate AI summary using OpenAI API.

        Args:
            text: Text to summarize
            prompt_template: Custom prompt (uses default if None).
                            Use {text} placeholder for input text.

        Returns:
            Generated summary text

        Raises:
            ValueError: If OPENAI_API_KEY environment variable not set
            ExecError: If API call fails

        Note:
            Requires OPENAI_API_KEY environment variable to be set in the container.
            Pass it via env parameter when creating DocGeneratorBox:
                async with DocGeneratorBox(env=[("OPENAI_API_KEY", key)]) as docgen:
        """
        default_prompt = "Analyze the following data and provide a concise summary:\\n\\n{text}\\n\\nSummary:"

        template = prompt_template or default_prompt
        prompt = template.replace("{text}", text[:1000])  # Limit to first 1000 chars

        script = """
import os
from openai import OpenAI

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("OPENAI_API_KEY environment variable not set")

prompt = data['prompt']

client = OpenAI(api_key=api_key)

response = client.chat.completions.create(
    model="gpt-4",
    messages=[{"role": "user", "content": prompt}]
)

print(response.choices[0].message.content)
"""

        data = {"prompt": prompt}

        return await self._exec_python_with_data(script, data)

    # ==================== File Management ====================

    async def upload_file(self, local_path: str, remote_path: str) -> str:
        """
        Upload file from host to container.

        Args:
            local_path: Local file path on host
            remote_path: Destination path in container

        Returns:
            Success message

        Raises:
            FileNotFoundError: If local file doesn't exist
            ValueError: If remote_path contains unsafe components (.. or absolute paths)
            ExecError: If upload fails

        Example:
            >>> await docgen.upload_file("./report.pptx", "/workspace/report.pptx")
        """
        import os

        # Validate remote path (container side)
        self._validate_safe_path(remote_path, allow_absolute=True)

        if not os.path.exists(local_path):
            raise FileNotFoundError(f"Local file not found: {local_path}")

        # Read file content
        with open(local_path, 'rb') as f:
            content = f.read()

        # Encode as base64
        b64_content = base64.b64encode(content).decode()

        script = """
import base64
import os

remote_path = data['remote_path']
b64_content = data['content']

# Create directory if needed (guard against basename-only paths)
dir_path = os.path.dirname(remote_path)
if dir_path:
    os.makedirs(dir_path, exist_ok=True)

# Decode and write file
content = base64.b64decode(b64_content)
with open(remote_path, 'wb') as f:
    f.write(content)

print(f"Uploaded to {remote_path} ({len(content)} bytes)")
"""

        data = {
            "remote_path": remote_path,
            "content": b64_content
        }

        return await self._exec_python_with_data(script, data)

    async def download_file(self, remote_path: str, local_path: str) -> str:
        """
        Download file from container to host.

        Args:
            remote_path: File path in container
            local_path: Destination path on host

        Returns:
            Success message

        Raises:
            ValueError: If paths contain unsafe components
            ExecError: If file doesn't exist or download fails

        Example:
            >>> await docgen.download_file("/workspace/report.pptx", "./report.pptx")
        """
        import os

        # Validate remote path (container side)
        self._validate_safe_path(remote_path, allow_absolute=True)

        # Validate local path (host side) - reject .. traversal
        abs_local_path = os.path.abspath(local_path)
        if '..' in os.path.normpath(local_path).split(os.sep):
            raise ValueError(f"Path traversal detected in local_path: {local_path}")

        script = """
import base64
import os

remote_path = data['remote_path']

if not os.path.exists(remote_path):
    raise FileNotFoundError(f"File not found: {remote_path}")

with open(remote_path, 'rb') as f:
    content = f.read()

# Encode as base64 for safe transmission
b64_content = base64.b64encode(content).decode()

import json
print(json.dumps({"content": b64_content, "size": len(content)}))
"""

        data = {"remote_path": remote_path}
        result = await self._exec_python_with_data(script, data)

        # Decode result with error handling
        try:
            result_data = json.loads(result)
        except json.JSONDecodeError as e:
            raise ExecError(f"Failed to decode container response: {e}\nOutput: {result[:200]}")

        content = base64.b64decode(result_data["content"])

        # Write to local file (guard against basename-only paths)
        local_dir = os.path.dirname(abs_local_path)
        if local_dir:
            os.makedirs(local_dir, exist_ok=True)
        with open(abs_local_path, 'wb') as f:
            f.write(content)

        return f"Downloaded to {local_path} ({result_data['size']} bytes)"

    async def list_files(self, directory: str = "/workspace", pattern: str = "*") -> List[Dict[str, Any]]:
        """
        List files in container directory.

        Args:
            directory: Directory to list
            pattern: Glob pattern (e.g., "*.pptx", "**/*.xlsx")

        Returns:
            List of file info dicts with keys: name, path, size, modified

        Example:
            >>> files = await docgen.list_files("/workspace", "*.pptx")
            >>> for f in files:
            ...     print(f"{f['name']}: {f['size']} bytes")
        """
        script = """
import os
import glob
from pathlib import Path
from datetime import datetime

directory = data['directory']
pattern = data['pattern']

search_path = os.path.join(directory, pattern)
files = []

for filepath in glob.glob(search_path, recursive=True):
    if os.path.isfile(filepath):
        stat = os.stat(filepath)
        files.append({
            "name": os.path.basename(filepath),
            "path": filepath,
            "size": stat.st_size,
            "modified": datetime.fromtimestamp(stat.st_mtime).isoformat()
        })

import json
print(json.dumps(files))
"""

        data = {"directory": directory, "pattern": pattern}
        result = await self._exec_python_with_data(script, data)
        return json.loads(result)

    async def delete_file(self, remote_path: str) -> str:
        """
        Delete file in container.

        Args:
            remote_path: File path to delete

        Returns:
            Success message

        Raises:
            ExecError: If file doesn't exist or deletion fails

        Example:
            >>> await docgen.delete_file("/workspace/old_report.pptx")
        """
        script = """
import os

remote_path = data['remote_path']

if not os.path.exists(remote_path):
    raise FileNotFoundError(f"File not found: {remote_path}")

os.remove(remote_path)
print(f"Deleted: {remote_path}")
"""

        data = {"remote_path": remote_path}
        return await self._exec_python_with_data(script, data)

    # ==================== Read Operations ====================

    async def read_ppt(self, file_path: str) -> Dict[str, Any]:
        """
        Read PowerPoint file and extract content.

        Args:
            file_path: Path to PPT file in container

        Returns:
            Dictionary with structure:
            {
                "slide_count": int,
                "slides": [
                    {
                        "slide_id": int,
                        "title": str,
                        "text_content": [str],
                        "notes": str
                    }
                ]
            }

        Example:
            >>> content = await docgen.read_ppt("/workspace/report.pptx")
            >>> print(f"Total slides: {content['slide_count']}")
            >>> for slide in content['slides']:
            ...     print(f"Slide {slide['slide_id']}: {slide['title']}")
        """
        script = """
from pptx import Presentation

file_path = data['file_path']
prs = Presentation(file_path)

slides = []
for i, slide in enumerate(prs.slides):
    slide_data = {
        "slide_id": i,
        "title": "",
        "text_content": [],
        "notes": ""
    }

    # Extract title
    if slide.shapes.title:
        slide_data["title"] = slide.shapes.title.text

    # Extract all text
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text:
            slide_data["text_content"].append(shape.text)

    # Extract notes
    if slide.has_notes_slide:
        notes_frame = slide.notes_slide.notes_text_frame
        if notes_frame:
            slide_data["notes"] = notes_frame.text

    slides.append(slide_data)

result = {
    "slide_count": len(slides),
    "slides": slides
}

import json
print(json.dumps(result, ensure_ascii=False))
"""

        data = {"file_path": file_path}
        result = await self._exec_python_with_data(script, data)
        return json.loads(result)

    async def read_word(self, file_path: str) -> Dict[str, Any]:
        """
        Read Word document and extract content.

        Args:
            file_path: Path to DOCX file in container

        Returns:
            Dictionary with structure:
            {
                "paragraphs": [str],
                "tables": [[[cell_text]]],
                "headings": [{"level": int, "text": str}]
            }

        Example:
            >>> content = await docgen.read_word("/workspace/report.docx")
            >>> print("\\n".join(content['paragraphs'][:3]))
        """
        script = """
from docx import Document

file_path = data['file_path']
doc = Document(file_path)

paragraphs = []
headings = []
tables = []

# Extract paragraphs and headings
for para in doc.paragraphs:
    if para.text.strip():
        paragraphs.append(para.text)
        if para.style.name.startswith('Heading'):
            level = int(para.style.name.split()[-1]) if para.style.name.split()[-1].isdigit() else 1
            headings.append({"level": level, "text": para.text})

# Extract tables
for table in doc.tables:
    table_data = []
    for row in table.rows:
        row_data = [cell.text for cell in row.cells]
        table_data.append(row_data)
    tables.append(table_data)

result = {
    "paragraph_count": len(paragraphs),
    "paragraphs": paragraphs,
    "table_count": len(tables),
    "tables": tables,
    "headings": headings
}

import json
print(json.dumps(result, ensure_ascii=False))
"""

        data = {"file_path": file_path}
        result = await self._exec_python_with_data(script, data)
        return json.loads(result)

    async def read_excel(self, file_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Read Excel file and extract data.

        Args:
            file_path: Path to XLSX file in container
            sheet_name: Specific sheet to read (None = all sheets)

        Returns:
            Dictionary with structure:
            {
                "sheet_names": [str],
                "sheets": {
                    "SheetName": [
                        {"col1": value, "col2": value, ...}
                    ]
                }
            }

        Example:
            >>> data = await docgen.read_excel("/workspace/data.xlsx", "Sheet1")
            >>> print(f"Rows: {len(data['sheets']['Sheet1'])}")
        """
        script = """
import pandas as pd

file_path = data['file_path']
sheet_name = data.get('sheet_name')

# Read Excel
if sheet_name:
    df_dict = {sheet_name: pd.read_excel(file_path, sheet_name=sheet_name)}
else:
    df_dict = pd.read_excel(file_path, sheet_name=None)

# Convert to dict
sheets = {}
for name, df in df_dict.items():
    sheets[name] = df.to_dict('records')

result = {
    "sheet_names": list(df_dict.keys()),
    "sheets": sheets
}

import json
print(json.dumps(result, ensure_ascii=False))
"""

        data = {"file_path": file_path, "sheet_name": sheet_name}
        result = await self._exec_python_with_data(script, data)
        return json.loads(result)

    # ==================== Modify Operations ====================

    async def modify_ppt(
            self,
            input_path: str,
            output_path: str,
            operations: List[Dict[str, Any]]
    ) -> str:
        """
        Modify existing PowerPoint file.

        Args:
            input_path: Path to input PPT in container
            output_path: Path to save modified PPT
            operations: List of modification operations

        Operations format:
            [
                {"op": "update_slide", "slide_id": 0, "title": "New Title", "content": {...}},
                {"op": "delete_slide", "slide_id": 1},
                {"op": "append_slide", "title": "...", "content": {...}}
            ]

        Returns:
            Success message

        Example:
            >>> await docgen.modify_ppt(
            ...     "/workspace/old.pptx",
            ...     "/workspace/new.pptx",
            ...     [{"op": "update_slide", "slide_id": 0, "title": "Updated"}]
            ... )
        """
        script = """
from pptx import Presentation
from pptx.util import Inches, Pt

input_path = data['input_path']
output_path = data['output_path']
operations = data['operations']

prs = Presentation(input_path)

for op in operations:
    if op['op'] == 'update_slide':
        slide_id = op['slide_id']
        if slide_id < len(prs.slides):
            slide = prs.slides[slide_id]

            # Update title
            if 'title' in op and slide.shapes.title:
                slide.shapes.title.text = op['title']

            # Update content (simplified)
            if 'content' in op and 'text' in op['content']:
                # Find first text box or create one
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame"):
                        shape.text_frame.text = op['content']['text']
                        break

    elif op['op'] == 'delete_slide':
        slide_id = op['slide_id']
        if 0 <= slide_id < len(prs.slides):
            rId = prs.slides._sldIdLst[slide_id].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[slide_id]

    elif op['op'] == 'append_slide':
        blank_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(blank_layout)

        # Add title
        if 'title' in op:
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
            title_frame = title_box.text_frame
            title_frame.text = op['title']
            title_frame.paragraphs[0].font.size = Pt(32)
            title_frame.paragraphs[0].font.bold = True

        # Add content
        if 'content' in op and 'text' in op['content']:
            text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
            text_frame = text_box.text_frame
            text_frame.text = op['content']['text']

prs.save(output_path)
print(f"Modified PPT saved to {output_path}")
"""

        data = {
            "input_path": input_path,
            "output_path": output_path,
            "operations": operations
        }

        return await self._exec_python_with_data(script, data)

    async def modify_word(
            self,
            input_path: str,
            output_path: str,
            operations: List[Dict[str, Any]]
    ) -> str:
        """
        Modify existing Word document.

        Args:
            input_path: Path to input DOCX in container
            output_path: Path to save modified DOCX
            operations: List of modification operations

        Operations format:
            [
                {"op": "replace_text", "find": "old", "replace": "new"},
                {"op": "append_paragraph", "text": "New paragraph"},
                {"op": "append_section", "heading": "...", "paragraphs": [...]}
            ]

        Returns:
            Success message

        Example:
            >>> await docgen.modify_word(
            ...     "/workspace/old.docx",
            ...     "/workspace/new.docx",
            ...     [{"op": "replace_text", "find": "Q1", "replace": "Q2"}]
            ... )
        """
        script = """
from docx import Document
from docx.shared import Pt

input_path = data['input_path']
output_path = data['output_path']
operations = data['operations']

doc = Document(input_path)

for op in operations:
    if op['op'] == 'replace_text':
        find_text = op['find']
        replace_text = op['replace']

        for para in doc.paragraphs:
            if find_text in para.text:
                para.text = para.text.replace(find_text, replace_text)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if find_text in cell.text:
                        cell.text = cell.text.replace(find_text, replace_text)

    elif op['op'] == 'append_paragraph':
        doc.add_paragraph(op['text'])

    elif op['op'] == 'append_section':
        if 'heading' in op:
            doc.add_heading(op['heading'], level=1)

        if 'paragraphs' in op:
            for para_text in op['paragraphs']:
                doc.add_paragraph(para_text)

doc.save(output_path)
print(f"Modified Word document saved to {output_path}")
"""

        data = {
            "input_path": input_path,
            "output_path": output_path,
            "operations": operations
        }

        return await self._exec_python_with_data(script, data)

    async def modify_excel(
            self,
            input_path: str,
            output_path: str,
            operations: List[Dict[str, Any]]
    ) -> str:
        """
        Modify existing Excel file.

        Args:
            input_path: Path to input XLSX in container
            output_path: Path to save modified XLSX
            operations: List of modification operations

        Operations format:
            [
                {"op": "update_cell", "sheet": "Sheet1", "cell": "A1", "value": 100},
                {"op": "append_row", "sheet": "Sheet1", "data": {"col1": 1, "col2": 2}},
                {"op": "delete_row", "sheet": "Sheet1", "row": 5}
            ]

        Returns:
            Success message

        Example:
            >>> await docgen.modify_excel(
            ...     "/workspace/old.xlsx",
            ...     "/workspace/new.xlsx",
            ...     [{"op": "update_cell", "sheet": "Sheet1", "cell": "A1", "value": 999}]
            ... )
        """
        script = """
import pandas as pd
from openpyxl import load_workbook

input_path = data['input_path']
output_path = data['output_path']
operations = data['operations']

# Load workbook
wb = load_workbook(input_path)

for op in operations:
    sheet_name = op.get('sheet', wb.sheetnames[0])
    ws = wb[sheet_name]

    if op['op'] == 'update_cell':
        cell = op['cell']
        value = op['value']
        ws[cell] = value

    elif op['op'] == 'delete_row':
        row = op['row']
        ws.delete_rows(row)

    elif op['op'] == 'append_row':
        # Use pandas for easier row append
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        new_row = pd.DataFrame([op['data']])
        df = pd.concat([df, new_row], ignore_index=True)

        # Write back
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sname in wb.sheetnames:
                if sname == sheet_name:
                    df.to_excel(writer, sheet_name=sname, index=False)
                else:
                    pd.read_excel(input_path, sheet_name=sname).to_excel(writer, sheet_name=sname, index=False)

        # Reload workbook
        wb = load_workbook(output_path)

wb.save(output_path)
print(f"Modified Excel saved to {output_path}")
"""

        data = {
            "input_path": input_path,
            "output_path": output_path,
            "operations": operations
        }

        return await self._exec_python_with_data(script, data)

    # ==================== Workspace Management ====================

    async def create_workspace(self, name: str, base_dir: str = "/workspace") -> str:
        """
        Create a new workspace directory.

        Args:
            name: Workspace name (must be a single path segment, no / or ..)
            base_dir: Base directory for workspaces

        Returns:
            Workspace path

        Raises:
            ValueError: If name contains unsafe components

        Example:
            >>> workspace = await docgen.create_workspace("project_alpha")
        """
        # Validate workspace name
        self._validate_workspace_name(name)

        script = """
import os

base_dir = data['base_dir']
name = data['name']

workspace_path = os.path.join(base_dir, name)
os.makedirs(workspace_path, exist_ok=True)

# Create subdirectories
os.makedirs(os.path.join(workspace_path, 'input'), exist_ok=True)
os.makedirs(os.path.join(workspace_path, 'output'), exist_ok=True)
os.makedirs(os.path.join(workspace_path, 'temp'), exist_ok=True)

print(workspace_path)
"""

        data = {"base_dir": base_dir, "name": name}
        return await self._exec_python_with_data(script, data)

    async def save_workspace(self, workspace_path: str, local_dir: str) -> str:
        """
        Download entire workspace to local directory.

        Args:
            workspace_path: Workspace path in container
            local_dir: Local directory to save workspace

        Returns:
            Success message with file count

        Raises:
            ValueError: If any file path attempts to escape local_dir

        Example:
            >>> await docgen.save_workspace("/workspace/project_alpha", "./workspaces/project_alpha")
        """
        import os

        # List all files in workspace
        files = await self.list_files(workspace_path, "**/*")

        # Download each file
        os.makedirs(local_dir, exist_ok=True)
        abs_local_dir = os.path.abspath(local_dir)

        for file_info in files:
            remote_file = file_info['path']
            # Calculate relative path
            rel_path = os.path.relpath(remote_file, workspace_path)
            local_file = os.path.join(abs_local_dir, rel_path)

            # Security check: ensure local_file is within local_dir
            abs_local_file = os.path.abspath(local_file)
            if not abs_local_file.startswith(abs_local_dir + os.sep):
                raise ValueError(f"Path traversal detected: {remote_file} -> {local_file}")

            # Download
            await self.download_file(remote_file, local_file)

        return f"Saved workspace: {len(files)} files to {local_dir}"

    async def load_workspace(self, workspace_path: str, local_dir: str) -> str:
        """
        Upload entire local directory to workspace.

        Args:
            workspace_path: Workspace path in container
            local_dir: Local directory to upload

        Returns:
            Success message with file count

        Example:
            >>> await docgen.load_workspace("/workspace/project_alpha", "./workspaces/project_alpha")
        """
        import os
        import glob

        # Find all files in local directory
        local_files = []
        for root, dirs, files in os.walk(local_dir):
            for file in files:
                local_path = os.path.join(root, file)
                rel_path = os.path.relpath(local_path, local_dir)
                remote_path = os.path.join(workspace_path, rel_path)
                local_files.append((local_path, remote_path))

        # Upload each file
        for local_path, remote_path in local_files:
            await self.upload_file(local_path, remote_path)

        return f"Loaded workspace: {len(local_files)} files to {workspace_path}"
