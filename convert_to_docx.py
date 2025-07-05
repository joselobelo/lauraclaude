import sys

try:
    import pypandoc
except ImportError:
    raise SystemExit("pypandoc is required. Install with `pip install pypandoc` and ensure pandoc is available.")


def convert(input_file: str, output_file: str) -> None:
    """Convert an input notebook or markdown file to docx with equations.

    This uses pandoc which preserves LaTeX equations as editable Word equations.
    """
    pypandoc.convert_file(
        input_file,
        "docx",
        outputfile=output_file,
        extra_args=["--mathml"],
    )


if __name__ == "__main__":
    input_path = sys.argv[1] if len(sys.argv) > 1 else "Laura.ipynb"
    output_path = sys.argv[2] if len(sys.argv) > 2 else "output.docx"
    convert(input_path, output_path)
    print(f"Archivo '{output_path}' generado correctamente.")
