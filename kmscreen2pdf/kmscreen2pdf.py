"""
kmscreen2pdf.exe

Usage:
  kmscreen2pdf.exe [--basedir=<dir>] [--filetype=<ext>]
  kmscreen2pdf.exe -h

Options:
  -h --help         Show this screen.
  --version         Show version.
  --basedir=<dir>   Select the base directory for the input and output.
  --filetype=<ext>  Select the file type for the image [default: wmf].

Description:
    This Windows program converts images and text files to PDFs with invisible text.
    Requires ImageMagick to be installed.
"""
import os
import re
import sys
import ctypes
import platform
import pathlib
from PIL import Image
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from loguru import logger
from docopt import docopt


class MessageBox:
    """Display a message box on Windows.
    Write the message to the log file as well."""

    # 0x00000000: OK button only.
    # 0x00000001: OK and Cancel buttons.
    # 0x00000010: Information icon.
    # 0x00000030: Warning icon.
    # 0x00000040: Error icon.

    window_types = {
        "info": 0x00000010,
        "warning": 0x00000030,
        "error": 0x00000040,
    }
    log_types = {
        "info": logger.info,
        "warning": logger.warning,
        "error": logger.error,
    }

    def _box(self, title: str, message: str, type: str):
        if platform.system() == "Windows":
            ok_only = 0x00000000
            selected_type = self.window_types.get(type, 0x00000040)
            ctypes.windll.user32.MessageBoxW(
                0, message, f"KMScreen2PDF: {title}", ok_only | selected_type
            )
        self.log_types.get(type, logger.error)(message)

    def error(self, title: str, message: str):
        self._box(title, message, "error")

    def warning(self, title: str, message: str):
        self._box(title, message, "warning")

    def info(self, title: str, message: str):
        self._box(title, message, "info")


window = MessageBox()
try:
    from wand.image import Image as WandImage
except ImportError:
    window.error("ImageMagick not found", "ImageMagick is required. Is it installed?")
    sys.exit(1)


def convert_img_to_png(img_path: pathlib.Path):
    """Convert image to PNG using wand (ImageMagick).
    For WMF files only Windows is supported."""
    png_path = img_path.with_suffix(".png")
    with WandImage(filename=img_path) as img:
        img.format = "png"
        img.save(filename=png_path)
    img_path.unlink()
    return png_path


def get_executable_path():
    "Finds the path of the python script or executable."
    if getattr(sys, "frozen", False):
        # The application is frozen
        return sys.executable
    # The application is not frozen
    return os.path.abspath(__file__)


def create_pdf_with_invisible_text(
    image_path, output_pdf_path, text, text_position=(50, 50), page_size=A4
):
    """Create a PDF with invisible text on top of an image."""
    # Create a canvas for the PDF
    c = canvas.Canvas(str(output_pdf_path), pagesize=page_size)

    # Convert WMF to PNG
    png_image_path = convert_img_to_png(image_path)

    # Open the image using Pillow
    img = Image.open(str(png_image_path))

    # Get image size
    img_width, img_height = img.size

    # Convert Pillow image to ReportLab-compatible format
    # temp_image = "temp_image.png"
    # img.save(temp_image)

    # Calculate position to center the image
    page_width, page_height = page_size
    x = (page_width - img_width // 1.5) // 2
    y = (page_height - img_height // 1.5) // 2

    # Draw the image on the PDF
    c.drawImage(png_image_path, x, y, width=img_width // 1.5, height=img_height // 1.5)

    # Set the text color to transparent
    transparent_color = Color(0, 0, 0, alpha=0)
    c.setFillColor(transparent_color)

    # Embed the invisible text on the PDF at the specified position
    # c.drawString(text_position[0], text_position[1], text)
    text_object = c.beginText()
    text_object.setTextOrigin(text_position[0], text_position[1])
    lines = re.split(r"\r\n|\r|\n", text)
    line_height = 14
    for i, line in enumerate(lines):
        text_object.setTextOrigin(text_position[0], text_position[1] - i * line_height)
        text_object.textLine(line)
    c.drawText(text_object)

    # Save the PDF
    c.showPage()
    c.save()

    # Delete the temporary PNG image
    # png_image_path.unlink()


def find_matching_documents(
    input_dir: pathlib.Path, image_extension: str = "jpg", text_extension: str = "txt"
):
    """Find matching image and text files in the input directory."""
    images = input_dir.glob(f"*.{image_extension}")
    for image in images:
        text_file = input_dir / image.with_suffix(f".{text_extension}")
        if text_file.exists():
            yield image, text_file
        else:
            window.warning(
                "Missing text file", f"Text file not found for image: '{image}'"
            )
            error_dir = input_dir.parent / "error"
            image.rename(error_dir / image.name)


def main():
    """Loop through the input directory,
    find matching image and text files,
    and create a PDF with invisible text."""
    arguments = docopt(__doc__, version="kmscreen2pdf 0.1.1")
    base_dir_from_arg = arguments.get("--basedir")
    if not base_dir_from_arg:
        base_dir = pathlib.Path(get_executable_path()).parent / "kmscreen2pdf"
    else:
        base_dir = pathlib.Path(base_dir_from_arg).expanduser().resolve()
    if not base_dir.exists():
        window.warning(
            "Base directory not found",
            f"Creating base directory: '{base_dir}' for input and output files.",
        )
        base_dir.mkdir(parents=True)
    input_dir = base_dir / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir = base_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    error_dir = base_dir / "error"
    error_dir.mkdir(parents=True, exist_ok=True)
    log_dir = base_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    logger.add(log_dir / "kmscreen2pdf.log", retention="30 days")
    for image_path, text_file_path in find_matching_documents(
        input_dir, image_extension=arguments["--filetype"]
    ):
        output_pdf_path = output_dir / image_path.with_suffix(".pdf").name
        png_image = image_path.with_suffix(".png")
        try:
            text = text_file_path.read_text(encoding="utf-8")
            text = text.replace("\n", "\r\n")
            create_pdf_with_invisible_text(image_path, output_pdf_path, text)
        except Exception as e:
            window.error(
                "Processing error",
                f"Error processing '{image_path}'. See log file for details.",
            )
            logger.critical(f"Error processing '{image_path}': {e}")
            text_file = image_path.with_suffix(".txt")
            image_path.rename(error_dir / image_path.name)
            text_file.rename(error_dir / text_file.name)
            if png_image.exists():
                png_image.rename(error_dir / png_image.name)
            raise e
        else:
            if image_path.exists():
                image_path.unlink()
            if text_file_path.exists():
                text_file_path.unlink()
            if png_image.exists():
                png_image.unlink()
    try:
        logger.debug("Processing complete.")
        logger.debug(f"The last image was '{image_path}'")
    except NameError:
        window.error(
            "No images found",
            f"No images and text files of the right type where found in the input directory: '{input_dir}'",
        )


if __name__ == "__main__":
    main()
