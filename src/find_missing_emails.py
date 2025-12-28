#!/usr/bin/env python3
"""
Missing Email Identifier

Scans an Excel file to identify rows with missing email addresses.
Reports the volume and details of records with missing emails.
"""

import sys
import re
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

import click
import pandas as pd


def is_valid_email(email: Any) -> bool:
    """
    Check if the value is a valid email address.

    Args:
        email: Value to check

    Returns:
        True if valid email, False otherwise
    """
    if pd.isna(email) or email is None:
        return False

    email_str = str(email).strip()
    if not email_str or email_str.lower() in ['nan', 'none', 'null', '', '-', 'n/a', 'na']:
        return False

    # Basic email regex pattern
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(email_pattern, email_str))


def find_email_column(df: pd.DataFrame) -> Optional[str]:
    """
    Auto-detect the email column in the dataframe.

    Args:
        df: DataFrame to search

    Returns:
        Column name if found, None otherwise
    """
    email_keywords = ['email', 'e-mail', 'mail', 'email_id', 'emailid', 'email_address',
                      'emailaddress', 'contact_email', 'author_email', 'corresponding_email']

    for col in df.columns:
        col_lower = str(col).lower().strip()
        for keyword in email_keywords:
            if keyword in col_lower:
                return col

    # Fallback: check column content for email patterns
    for col in df.columns:
        sample = df[col].dropna().head(10)
        email_count = sum(1 for val in sample if is_valid_email(val))
        if email_count >= 3:  # At least 3 valid emails in sample
            return col

    return None


def find_identifier_columns(df: pd.DataFrame) -> List[str]:
    """
    Find columns that can help identify records (name, title, ID, etc.).

    Args:
        df: DataFrame to search

    Returns:
        List of column names useful for identification
    """
    identifier_keywords = ['name', 'title', 'id', 'author', 'article', 'volume', 'issue',
                          'paper', 'doi', 'serial', 'number', 'index', 'row']

    identifier_cols = []
    for col in df.columns:
        col_lower = str(col).lower().strip()
        for keyword in identifier_keywords:
            if keyword in col_lower:
                identifier_cols.append(col)
                break

    return identifier_cols[:5]  # Limit to 5 most relevant columns


def find_name_columns(df: pd.DataFrame) -> List[str]:
    """
    Find author name columns for deduplication.

    Args:
        df: DataFrame to search

    Returns:
        List of name column names (first, middle, last)
    """
    name_cols = []
    for col in df.columns:
        col_lower = str(col).lower()
        if any(x in col_lower for x in ['fname', 'first_name', 'firstname', 'lname', 'last_name', 'lastname']):
            name_cols.append(col)
    return name_cols


def get_cell_value(row, col):
    """Safely get a cell value, handling Series from duplicate columns."""
    val = row[col]
    if isinstance(val, pd.Series):
        val = val.iloc[0]
    return val if pd.notna(val) else ""


def deduplicate_by_author(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """
    Remove duplicate authors from the missing emails dataframe.

    Args:
        df: DataFrame with missing emails

    Returns:
        Tuple of (deduplicated DataFrame, count of duplicates removed)
    """
    name_cols = find_name_columns(df)

    if not name_cols:
        # Fallback: look for any column with 'author' and 'name' in it
        for col in df.columns:
            col_lower = str(col).lower()
            if 'author' in col_lower and ('name' in col_lower or 'fname' in col_lower or 'lname' in col_lower):
                name_cols.append(col)

    if not name_cols:
        return df, 0

    # Create a combined name key for deduplication
    def make_name_key(row):
        parts = []
        for col in name_cols:
            val = get_cell_value(row, col)
            if val:
                parts.append(str(val).strip().lower())
        return " ".join(sorted(parts))

    df = df.copy()
    df['_name_key'] = df.apply(make_name_key, axis=1)

    original_count = len(df)
    df_deduped = df.drop_duplicates(subset=['_name_key'], keep='first')
    df_deduped = df_deduped.drop(columns=['_name_key'])

    duplicates_removed = original_count - len(df_deduped)

    return df_deduped, duplicates_removed


def analyze_missing_emails(
    file_path: Path,
    email_column: Optional[str] = None,
    sheet_name: Optional[str] = None,
    volume: Optional[int] = None,
    issue: Optional[int] = None
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Analyze an Excel file for missing email addresses.

    Args:
        file_path: Path to the Excel file
        email_column: Name of the email column (auto-detected if not provided)
        sheet_name: Name of the sheet to analyze (first sheet if not provided)
        volume: Optional volume number to filter by
        issue: Optional issue number to filter by

    Returns:
        Tuple of (DataFrame with missing emails, statistics dict)
    """
    # Read the Excel file
    if sheet_name:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    else:
        df = pd.read_excel(file_path)

    # Auto-detect email column if not provided
    if not email_column:
        email_column = find_email_column(df)
        if not email_column:
            raise ValueError(
                "Could not auto-detect email column. Please specify with --email-column. "
                f"Available columns: {list(df.columns)}"
            )

    if email_column not in df.columns:
        raise ValueError(
            f"Column '{email_column}' not found. Available columns: {list(df.columns)}"
        )

    # Filter by volume/issue if specified
    original_count = len(df)
    if volume is not None:
        vol_cols = [c for c in df.columns if 'vol' in str(c).lower()]
        if vol_cols:
            df = df[df[vol_cols[0]] == volume]

    if issue is not None:
        iss_cols = [c for c in df.columns if 'iss' in str(c).lower()]
        if iss_cols:
            df = df[df[iss_cols[0]] == issue]

    filtered_count = len(df)

    # Find rows with missing emails
    df['_has_valid_email'] = df[email_column].apply(is_valid_email)
    missing_emails_df = df[~df['_has_valid_email']].copy()
    missing_emails_df = missing_emails_df.drop(columns=['_has_valid_email'])

    # Add row numbers (1-indexed for user readability)
    missing_emails_df['Excel_Row'] = missing_emails_df.index + 2  # +2 for header and 0-index

    # Calculate statistics
    total_records = filtered_count
    missing_count = len(missing_emails_df)
    valid_count = total_records - missing_count
    missing_percentage = (missing_count / total_records * 100) if total_records > 0 else 0

    stats = {
        'file_path': str(file_path),
        'email_column': email_column,
        'total_records': total_records,
        'original_records': original_count,
        'filtered_records': filtered_count,
        'missing_emails': missing_count,
        'valid_emails': valid_count,
        'missing_percentage': round(missing_percentage, 2),
        'volume_filter': volume,
        'issue_filter': issue,
    }

    return missing_emails_df, stats


def print_report(
    missing_df: pd.DataFrame,
    stats: Dict[str, Any],
    email_column: str,
    show_details: bool = True,
    max_rows: int = 50
):
    """
    Print a formatted report of missing emails.

    Args:
        missing_df: DataFrame containing rows with missing emails
        stats: Statistics dictionary
        email_column: Name of the email column
        show_details: Whether to show individual record details
        max_rows: Maximum rows to display in details
    """
    click.echo("\n" + "=" * 70)
    click.echo("           MISSING EMAIL ANALYSIS REPORT")
    click.echo("=" * 70)

    # File info
    click.echo(f"\nFile: {stats['file_path']}")
    click.echo(f"Email Column: {stats['email_column']}")

    # Filters applied
    if stats['volume_filter'] or stats['issue_filter']:
        click.echo("\nFilters Applied:")
        if stats['volume_filter']:
            click.echo(f"  - Volume: {stats['volume_filter']}")
        if stats['issue_filter']:
            click.echo(f"  - Issue: {stats['issue_filter']}")
        click.echo(f"  - Records after filtering: {stats['filtered_records']} (from {stats['original_records']})")

    # Summary statistics
    click.echo("\n" + "-" * 70)
    click.echo("                      SUMMARY")
    click.echo("-" * 70)
    click.echo(f"  Total Records:          {stats['total_records']}")
    click.echo(f"  Valid Emails:           {stats['valid_emails']}")
    click.echo(f"  Missing Emails (raw):   {stats['missing_emails']}")
    if 'duplicates_removed' in stats:
        click.echo(f"  Duplicate Authors:      {stats['duplicates_removed']}")
        click.echo(f"  Unique Missing:         {stats['unique_missing']}")
    click.echo(f"  Missing Percentage:     {stats['missing_percentage']}%")
    click.echo("-" * 70)

    # Issue severity
    unique_count = stats.get('unique_missing', stats['missing_emails'])
    if stats['missing_percentage'] == 0:
        click.echo("\nStatus: All records have valid email addresses!")
    elif stats['missing_percentage'] < 5:
        click.echo(f"\nStatus: LOW - Only {unique_count} unique authors missing emails")
    elif stats['missing_percentage'] < 20:
        click.echo(f"\nStatus: MODERATE - {unique_count} unique authors need attention")
    else:
        click.echo(f"\nStatus: HIGH - {unique_count} unique authors require email collection")

    # Detailed listing
    if show_details and len(missing_df) > 0:
        click.echo("\n" + "-" * 70)
        click.echo("            UNIQUE AUTHORS WITH MISSING EMAILS")
        click.echo("-" * 70)

        # Find name columns for cleaner display
        name_cols = find_name_columns(missing_df)

        # Build clean display
        display_df = missing_df.head(max_rows).copy()

        click.echo("")
        for idx, (_, row) in enumerate(display_df.iterrows(), 1):
            # Get author name
            name_parts = []
            for col in name_cols:
                val = get_cell_value(row, col)
                if val:
                    name_parts.append(str(val).strip())

            author_name = " ".join(name_parts) if name_parts else "(Unknown)"

            # Get title if available
            title = ""
            for col in missing_df.columns:
                if 'title' in str(col).lower():
                    title = str(get_cell_value(row, col))[:50]
                    break

            # Get volume if available
            volume = ""
            for col in missing_df.columns:
                if 'vol' in str(col).lower():
                    val = get_cell_value(row, col)
                    if val:
                        volume = str(val)
                    break

            # Get issue if available
            issue = ""
            for col in missing_df.columns:
                if 'iss' in str(col).lower():
                    val = get_cell_value(row, col)
                    if val:
                        issue = str(val)
                    break

            excel_row = int(row['Excel_Row']) if 'Excel_Row' in row else "?"

            click.echo(f"  {idx:>3}. {author_name}")
            if volume or issue:
                vol_iss_parts = []
                if volume:
                    vol_iss_parts.append(f"Vol: {volume}")
                if issue:
                    vol_iss_parts.append(f"Issue: {issue}")
                click.echo(f"       {', '.join(vol_iss_parts)}")
            if title:
                click.echo(f"       Article: {title}...")
            click.echo(f"       Excel Row: {excel_row}")
            click.echo("")

        if len(missing_df) > max_rows:
            click.echo(f"  ... and {len(missing_df) - max_rows} more authors\n")

    click.echo("=" * 70)


def export_missing_emails(
    missing_df: pd.DataFrame,
    output_path: Path,
    stats: Dict[str, Any]
):
    """
    Export missing emails report to Excel.

    Args:
        missing_df: DataFrame with missing email records
        output_path: Path for output file
        stats: Statistics dictionary
    """
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Write missing records
        missing_df.to_excel(writer, sheet_name='Missing_Emails', index=False)

        # Write summary
        summary_df = pd.DataFrame([stats])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

    click.echo(f"\nExported report to: {output_path}")


@click.command()
@click.argument('excel_file', type=click.Path(exists=True, path_type=Path))
@click.option('--email-column', '-e', help='Name of the email column (auto-detected if not specified)')
@click.option('--sheet', '-s', help='Sheet name to analyze (first sheet if not specified)')
@click.option('--volume', '-v', type=int, help='Filter by volume number')
@click.option('--issue', '-i', type=int, help='Filter by issue number')
@click.option('--export', '-o', type=click.Path(path_type=Path), help='Export results to Excel file')
@click.option('--no-details', is_flag=True, help='Hide individual record details')
@click.option('--max-rows', default=50, type=int, help='Maximum rows to display (default: 50)')
def main(
    excel_file: Path,
    email_column: Optional[str],
    sheet: Optional[str],
    volume: Optional[int],
    issue: Optional[int],
    export: Optional[Path],
    no_details: bool,
    max_rows: int
):
    """
    Identify missing email addresses in an Excel file.

    Analyzes the specified Excel file and reports:
    - Total count of records with missing emails
    - Percentage of records affected
    - Details of each record with a missing email

    Example:
        python find_missing_emails.py data.xlsx
        python find_missing_emails.py data.xlsx -v 18 -i 1 --export missing_report.xlsx
    """
    try:
        click.echo(f"Analyzing: {excel_file}")

        # Analyze the file
        missing_df, stats = analyze_missing_emails(
            file_path=excel_file,
            email_column=email_column,
            sheet_name=sheet,
            volume=volume,
            issue=issue
        )

        # Deduplicate by author name
        missing_df_unique, duplicates_removed = deduplicate_by_author(missing_df)
        stats['duplicates_removed'] = duplicates_removed
        stats['unique_missing'] = len(missing_df_unique)

        # Print report with deduplicated data
        print_report(
            missing_df=missing_df_unique,
            stats=stats,
            email_column=stats['email_column'],
            show_details=not no_details,
            max_rows=max_rows
        )

        # Export if requested
        if export:
            export_missing_emails(missing_df_unique, export, stats)

        # Exit with error code if missing emails found
        if stats['missing_emails'] > 0:
            sys.exit(1)

    except FileNotFoundError:
        click.echo(f"Error: File not found: {excel_file}", err=True)
        sys.exit(1)
    except ValueError as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"Unexpected error: {e}", err=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
