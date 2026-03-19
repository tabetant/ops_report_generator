from dotenv import load_dotenv
import os

load_dotenv()
api_key = os.getenv("ANTHROPIC_API_KEY")

import typer
import loader
import analyzer
import output

app = typer.Typer(help="Generate operational reports from CSV or Excel data files.")


@app.command()
def generate(
    file_path: str = typer.Argument(..., help="Path to a CSV or Excel (.xlsx) file"),
    mock: bool = typer.Option(True, "--mock/--no-mock", help="Use mock Claude response (default: True)"),
    output_path: str = typer.Option("report.docx", "--output", help="Path for the generated Word report"),
    model: str = typer.Option("claude-haiku-4-5-20251001", "--model", help="Claude model to use for real API calls"),
):
    """Load operational data, analyze it with Claude, and produce a terminal + Word report."""

    if not mock and not api_key:
        typer.echo(
            "Error: ANTHROPIC_API_KEY is not set.\n"
            "Create a .env file in the project root with:\n"
            "  ANTHROPIC_API_KEY=your-key-here",
            err=True,
        )
        raise typer.Exit(code=1)

    try:
        df = loader.load_file(file_path)
        loader.validate_columns(df, loader.REQUIRED_COLUMNS)
        summary = loader.summarize(df)

        prompt = analyzer.build_prompt(summary)
        analysis = analyzer.call_claude(prompt, mock=mock)

        output.print_terminal_report(summary, analysis)
        saved_path = output.generate_word_report(summary, analysis, output_path)
        typer.echo(f"Word report saved to: {saved_path}")

    except FileNotFoundError as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(code=1)
    except ValueError as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(code=1)
    except EnvironmentError as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(code=1)
    except Exception as e:
        name = type(e).__name__
        typer.echo(f"Error ({name}): {e}", err=True)
        raise typer.Exit(code=1)


if __name__ == "__main__":
    app()
