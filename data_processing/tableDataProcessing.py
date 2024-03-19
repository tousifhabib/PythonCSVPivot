def preprocess_dataframe(df):
    """Prepare the dataframe by copying and filling NaN values."""
    df = df.copy()
    df.fillna("", inplace=True)
    return df


def assign_subtotal_levels(df, dynamic_cols):
    """Assign levels to subtotals and grand totals in the dataframe."""
    df['Subtotal_Level'] = None
    assign_subtotal_for_dynamic_cols(df, dynamic_cols)
    mark_grand_totals(df, dynamic_cols)


def assign_subtotal_for_dynamic_cols(df, dynamic_cols):
    """Assign subtotal levels based on dynamic columns."""
    for idx, (current_col, next_col) in enumerate(zip(dynamic_cols[:-1], dynamic_cols[1:])):
        is_subtotal = df[current_col].eq("") & df[next_col].ne("")
        df.loc[is_subtotal, 'Subtotal_Level'] = idx + 1


def mark_grand_totals(df, dynamic_cols):
    """Mark the grand total rows specifically."""
    df.loc[df[dynamic_cols[-1]] == "Grand Total", 'Subtotal_Level'] = 0
    subtotal_and_total_rows = df[dynamic_cols[-1]].isin(["", "Grand Total"])
    df.loc[subtotal_and_total_rows, dynamic_cols[:-1]] = ""


def clean_up_columns(df, dynamic_cols):
    """Clean up the columns to ensure proper formatting for the table."""
    for idx, col in enumerate(dynamic_cols[:-1]):
        df[col] = df[col].where(~(
                df[col].eq(df[col].shift()) &
                df[dynamic_cols[:idx]].eq(df[dynamic_cols[:idx]].shift(1)).all(axis=1)
        ), "")


def format_table_data(df):
    """Format the dataframe into a list of lists for table data."""
    table_data = [df.columns.tolist()] + df.values.tolist()
    for row in table_data:
        del row[-1]
    return table_data
