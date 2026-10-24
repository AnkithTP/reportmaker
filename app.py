
import csv
import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)
from openpyxl.utils import get_column_letter


# ==================================================
# 1. PAGE CONFIGURATION
# ==================================================

st.set_page_config(
    page_title="PostgreSQL CSV to Excel Converter",
    page_icon="📊",
    layout="wide",
)

st.title("PostgreSQL CSV to Excel Converter")
st.write(
    "Upload your PostgreSQL report CSV to generate "
    "a formatted Excel workbook automatically."
)


# ==================================================
# 2. FIND AND EXTRACT CSV SECTIONS
# ==================================================

def find_section(lines, heading, expected_header):
    """
    Find a section by its title and expected column header.
    Supports different delimiters in different sections.
    """

    start = next(
        (
            i for i, line in enumerate(lines)
            if line.strip().lower() == heading.lower()
        ),
        None,
    )

    if start is None:
        raise ValueError(
            f"Section heading not found: {heading}"
        )

    header_index = next(
        (
            i for i in range(start + 1, len(lines))

            
            if line_matches(lines[i], expected_header)
        ),
        None,
    )

    if header_index is None:
        raise ValueError(
            f"Expected header not found: {expected_header}"
        )

    known_headings = {
        "apollo prod last vacuum date",
        "apollo prod all database size",
        "apollo prod all tables size",
    }

    end = len(lines)

    for i in range(header_index + 1, len(lines)):
        current_line = lines[i].strip()

        if re.fullmatch(r"\(\d+\s+rows?\)", current_line):
            end = i
            break

        if current_line.lower() in known_headings:
            end = i
            break

    section_lines = [
        line.strip()
        for line in lines[header_index:end]
        if line.strip()
        and not re.fullmatch(
            r"\(\d+\s+rows?\)",
            line.strip(),
        )
    ]

    if len(section_lines) < 2:
        raise ValueError(
            f"No data rows found in section: {heading}"
        )

    return section_lines


def line_matches(line, expected_header):
    """Match a header while ignoring surrounding whitespace."""
    return line.strip().lstrip("\ufeff").lower() == (
        expected_header.lower()
    )


# ==================================================
# 3. PARSE EACH DELIMITED SECTION
# ==================================================

def parse_delimited_section(section_lines, delimiter):
    """Parse a CSV or pipe-delimited section into a DataFrame."""

    reader = csv.DictReader(
        io.StringIO("\n".join(section_lines)),
        delimiter=delimiter,
        skipinitialspace=True,
    )

    if not reader.fieldnames:
        raise ValueError("Could not read column headers.")

    reader.fieldnames = [
        name.strip().lstrip("\ufeff")
        for name in reader.fieldnames
    ]

    records = []

    for row in reader:
        cleaned = {
            key.strip(): (
                value.strip()
                if isinstance(value, str)
                else value
            )
            for key, value in row.items()
            if key is not None
        }

        if any(
            value not in (None, "")
            for value in cleaned.values()
        ):
            records.append(cleaned)

    return pd.DataFrame(records)


def validate_columns(df, required_columns, section_name):
    """Ensure all required source columns are available."""

    missing = set(required_columns) - set(df.columns)

    if missing:
        raise ValueError(
            f"{section_name} is missing columns: "
            f"{', '.join(sorted(missing))}"
        )


# ==================================================
# 4. PROCESS THE COMPLETE CSV
# ==================================================

def parse_report(uploaded_file):
    """
    Extract:
    - Vacuum statistics
    - Database sizes
    - Table sizes
    """

    try:
        content = uploaded_file.getvalue().decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError(
            "Unable to decode this CSV. Please save it as UTF-8."
        ) from exc

    lines = content.splitlines()

    # Vacuum statistics use commas.
    vacuum_lines = find_section(
        lines,
        "Apollo Prod last VACUUM date",
        "relname,n_live_tup,n_dead_tup,last_vacuum,last_autovacuum",
    )

    # Database sizes use pipes.
    database_lines = find_section(
        lines,
        "Apollo Prod all database size",
        "databasename|sizemb",
    )

    # Table sizes use pipes.
    table_lines = find_section(
        lines,
        "Apollo Prod all tables size",
        "tablename|pg_relation_size",
    )

    vacuum_df = parse_delimited_section(
        vacuum_lines, ","
    )

    database_df = parse_delimited_section(
        database_lines, "|"
    )

    table_df = parse_delimited_section(
        table_lines, "|"
    )

    # Validate source columns.
    validate_columns(
        vacuum_df,
        [
            "relname",
            "n_live_tup",
            "n_dead_tup",
            "last_vacuum",
            "last_autovacuum",
        ],
        "Vacuum statistics",
    )

    validate_columns(
        database_df,
        ["databasename", "sizemb"],
        "Database sizes",
    )

    validate_columns(
        table_df,
        ["tablename", "pg_relation_size"],
        "Table sizes",
    )

    # Validate numerical fields.
    for column in ["n_live_tup", "n_dead_tup"]:
        vacuum_df[column] = pd.to_numeric(
            vacuum_df[column],
            errors="raise",
        ).astype("Int64")

    database_df["sizemb"] = pd.to_numeric(
        database_df["sizemb"],
        errors="raise",
    )

    table_df["pg_relation_size"] = pd.to_numeric(
        table_df["pg_relation_size"],
        errors="raise",
    )

    # Convert table bytes to MB.
    table_df["pg_relation_size(MB)"] = (
        table_df["pg_relation_size"] / (1024 ** 2)
    ).round(3)

    # Convert database MB to GB.
    database_df["size(GB)"] = (
        database_df["sizemb"] / 1024
    ).round(3)

    # Preserve the required column order.
    vacuum_df = vacuum_df[
        [
            "relname",
            "n_live_tup",
            "n_dead_tup",
            "last_vacuum",
            "last_autovacuum",
        ]
    ]

    table_df = table_df[
        [
            "tablename",
            "pg_relation_size",
            "pg_relation_size(MB)",
        ]
    ]

    database_df = database_df[
        [
            "databasename",
            "sizemb",
            "size(GB)",
        ]
    ]

    return vacuum_df, table_df, database_df


# ==================================================
# 5. GENERATE FORMATTED EXCEL
# ==================================================

def generate_excel(vacuum_df, table_df, database_df):
    """Generate one Excel sheet containing three formatted tables."""

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        workbook = writer.book
        ws = workbook.create_sheet("Sheet1")

        sections = [
            (
                "Apollo Prod last VACUUM date",
                vacuum_df,
                0,
            ),
            (
                "Apollo Prod all tables size",
                table_df,
                7,
            ),
            (
                "Apollo Prod all database size",
                database_df,
                12,
            ),
        ]

        # Yellow section title.
        yellow_fill = PatternFill(
            fill_type="solid",
            fgColor="FFFF00",
        )

        # Light-yellow column header.
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="FFF2CC",
        )

        title_font = Font(
            bold=True,
            color="000000",
            size=11,
        )

        header_font = Font(
            bold=True,
            color="000000",
        )

        # Thin light-grey borders.
        thin_side = Side(
            style="thin",
            color="B7B7B7",
        )

        table_border = Border(
            left=thin_side,
            right=thin_side,
            top=thin_side,
            bottom=thin_side,
        )

        # ------------------------------------------
        # WRITE ALL THREE TABLES
        # ------------------------------------------

        for title, df, start_col in sections:

            first_col = start_col + 1
            last_col = start_col + len(df.columns)

            # Merge the section title across its columns.
            if last_col > first_col:
                ws.merge_cells(
                    start_row=1,
                    start_column=first_col,
                    end_row=1,
                    end_column=last_col,
                )

            title_cell = ws.cell(
                row=1,
                column=first_col,
            )

            title_cell.value = title
            title_cell.fill = yellow_fill
            title_cell.font = title_font
            title_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )
            title_cell.border = table_border

            # Fill every cell in the title row.
            for col in range(first_col, last_col + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = yellow_fill
                cell.border = table_border

            # Column headers in row 2.
            for col_offset, column_name in enumerate(df.columns):
                excel_col = first_col + col_offset

                cell = ws.cell(
                    row=2,
                    column=excel_col,
                )

                cell.value = column_name
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )
                cell.border = table_border

                # Write data starting from row 3.
                for row_offset, value in enumerate(
                    df[column_name].tolist(),
                    start=3,
                ):
                    data_cell = ws.cell(
                        row=row_offset,
                        column=excel_col,
                    )

                    if pd.isna(value):
                        data_cell.value = None
                    elif hasattr(value, "item"):
                        data_cell.value = value.item()
                    else:
                        data_cell.value = value

                    data_cell.border = table_border

        # ------------------------------------------
        # COLUMN WIDTHS
        # ------------------------------------------

        widths = {
            "A": 32,
            "B": 15,
            "C": 15,
            "D": 29,
            "E": 29,
            "H": 34,
            "I": 22,
            "J": 24,
            "M": 25,
            "N": 15,
            "O": 15,
        }

        for column, width in widths.items():
            ws.column_dimensions[column].width = width

        # ------------------------------------------
        # NUMBER FORMATTING
        # ------------------------------------------

        for row in range(3, ws.max_row + 1):

            # Vacuum counts remain whole numbers.
            ws.cell(row, 2).number_format = "#,##0"
            ws.cell(row, 3).number_format = "#,##0"

            # Table size in bytes remains a whole number.
            ws.cell(row, 9).number_format = "#,##0"

            # Table size in MB: three decimals.
            ws.cell(row, 10).number_format = "#,##0.000"

            # Database size in MB: three decimals.
            ws.cell(row, 14).number_format = "#,##0.000"

            # Database size in GB: three decimals.
            ws.cell(row, 15).number_format = "#,##0.000"

        # ------------------------------------------
        # WORKSHEET LAYOUT
        # ------------------------------------------

        ws.freeze_panes = "A3"
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 32

        # Separate autofilters for each table.
        for _, df, start_col in sections:

            first_col = start_col + 1
            last_col = start_col + len(df.columns)
            last_row = len(df) + 2

            if len(df) > 0:
                ws.auto_filter.ref = (
                    f"{get_column_letter(first_col)}2:"
                    f"{get_column_letter(last_col)}{last_row}"
                )

    output.seek(0)

    return output.getvalue()


# ==================================================
# 6. STREAMLIT USER INTERFACE
# ==================================================

uploaded_file = st.file_uploader(
    "Upload your PostgreSQL report CSV",
    type=["csv"],
    help=(
        "Upload the CSV containing vacuum statistics, "
        "database sizes and table sizes."
    ),
)

if uploaded_file is not None:

    try:
        vacuum_df, table_df, database_df = parse_report(
            uploaded_file
        )

        st.success("CSV processed successfully.")

        # Summary cards.
        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Vacuum records",
            len(vacuum_df),
        )

        col2.metric(
            "Table records",
            len(table_df),
        )

        col3.metric(
            "Database records",
            len(database_df),
        )

        # Data previews.
        with st.expander(
            "Preview vacuum statistics",
            expanded=True,
        ):
            st.dataframe(
                vacuum_df,
                use_container_width=True,
            )

        with st.expander("Preview table sizes"):
            st.dataframe(
                table_df,
                use_container_width=True,
            )

        with st.expander("Preview database sizes"):
            st.dataframe(
                database_df,
                use_container_width=True,
            )

        # Generate Excel workbook.
        excel_bytes = generate_excel(
            vacuum_df,
            table_df,
            database_df,
        )

        output_filename = (
            f"{Path(uploaded_file.name).stem}.xlsx"
        )

        st.download_button(
            label="Download Excel Report",
            data=excel_bytes,
            file_name=output_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
        )

    except Exception as exc:
        st.error(f"Unable to process the CSV: {exc}")

        st.info(
            "Verify that the uploaded CSV contains all three "
            "expected section titles and their column headers."
        )
