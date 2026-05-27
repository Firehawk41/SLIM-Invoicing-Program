import logging

import click
from sqlalchemy.orm import Session

from infrastructure.database import Base, make_engine
from slim.app import create_app
from slim.pipeline.sales_order_csv_writer import SalesOrderCsvWriter

_logger = logging.getLogger(__name__)


@click.group()
@click.option(
    "--db",
    default=None,
    show_default=True,
    help="SQLAlchemy database URL. Overrides DB_URL in .env and the built-in SQLite default.",
)
@click.option("--verbose", is_flag=True, default=False, help="Enable DEBUG logging.")
@click.pass_context
def cli(ctx: click.Context, db: str, verbose: bool) -> None:
    """LabPlus invoice pipeline."""
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.WARNING,
        format="%(levelname)s %(name)s: %(message)s",
    )
    ctx.ensure_object(dict)
    ctx.obj["db"] = db


@cli.command()
@click.option(
    "--folder",
    required=True,
    type=click.Path(exists=True, file_okay=False),
    help="Folder containing .xlsx invoice files.",
)
@click.option(
    "--start",
    required=True,
    type=click.DateTime(formats=["%Y-%m-%d"]),
    help="Start date inclusive (YYYY-MM-DD).",
)
@click.option(
    "--end",
    required=True,
    type=click.DateTime(formats=["%Y-%m-%d"]),
    help="End date inclusive (YYYY-MM-DD).",
)
@click.option(
    "--output",
    required=True,
    type=click.Path(dir_okay=False, writable=True),
    help="Output CSV path.",
)
@click.pass_context
def run(ctx: click.Context, folder: str, start, end, output: str) -> None:
    """Process a folder of invoice files between two dates."""
    start_date = start.date()
    end_date = end.date()
    db_url: str | None = ctx.obj["db"]

    engine = make_engine(db_url)
    Base.metadata.create_all(engine)

    with Session(engine) as session:
        app = create_app(session)
        submissions = app.loader.load_by_date_range(folder, start_date, end_date)
        written = 0
        with SalesOrderCsvWriter(output) as writer:
            for submission in submissions:
                try:
                    so = app.builder.build_from_submission(submission)
                    writer.write(so)
                    written += 1
                except Exception as exc:
                    _logger.error("Failed to build sales order for %s: %s", submission.file_name, exc)

    click.echo(f"Processed {len(submissions)} submission(s), wrote {written} sales order(s) to {output}.")


@cli.command(name="run-single")
@click.option(
    "--file",
    "file_path",
    required=True,
    type=click.Path(exists=True, dir_okay=False),
    help="Path to a single .xlsx invoice file.",
)
@click.option(
    "--output",
    required=True,
    type=click.Path(dir_okay=False, writable=True),
    help="Output CSV path.",
)
@click.pass_context
def run_single(ctx: click.Context, file_path: str, output: str) -> None:
    """Process a single invoice file."""
    db_url: str | None = ctx.obj["db"]

    engine = make_engine(db_url)
    Base.metadata.create_all(engine)

    try:
        with Session(engine) as session:
            app = create_app(session)
            submission = app.loader.load_single(file_path)
            so = app.builder.build_from_submission(submission)
            with SalesOrderCsvWriter(output) as writer:
                writer.write(so)
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc

    click.echo(f"Wrote 1 sales order to {output}.")
