import pandas as pd
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

def load_data(file_path):
    try:
        return pd.read_csv(file_path)
    except Exception as e:
        print(f"Error loading data: {e}")
        return pd.DataFrame()


def apply_filters(df, filters):
    for filter_col, filter_value in filters.items():
        df = df[df[filter_col] == filter_value]
    return df


def group_data(df, group_cols, agg_func):
    if agg_func == "count":
        return df.groupby(group_cols).size().reset_index(name="Count")
    else:
        return df.groupby(group_cols).agg(agg_func).reset_index()


def preprocess_data(df, filters, group_cols, agg_func):
    missing_filters = [f for f in filters if f not in df.columns]
    if missing_filters:
        raise ValueError(f"Filter columns {missing_filters} not found in data.")

    missing_groups = [g for g in group_cols if g not in df.columns]
    if missing_groups:
        raise ValueError(f"Group columns {missing_groups} not found in data.")

    filtered_data = apply_filters(df, filters)
    grouped_data = group_data(filtered_data, group_cols, agg_func)

    return grouped_data


def calculate_subtotals(grouped_data, group_by_col, agg_col):
    subtotals = (
        grouped_data.groupby(group_by_col)[agg_col].sum().reset_index(name=agg_col)
    )
    return subtotals


def add_grand_total(subtotals, subtotal_col, agg_col):
    grand_total_row = ["Grand Total"]
    grand_total_row.extend([""] * (len(subtotal_col) - 1))
    grand_total_row.append(subtotals[agg_col].sum())
    return pd.DataFrame([grand_total_row], columns=subtotal_col + [agg_col])


def prepare_subtotal_rows(grouped_data, subtotal_cols, agg_col):
    subtotals = calculate_subtotals(grouped_data, subtotal_cols, agg_col)
    grouped_data["key"] = grouped_data[subtotal_cols].apply(tuple, axis=1)
    subtotals["key"] = subtotals[subtotal_cols].apply(tuple, axis=1)

    final_data = pd.DataFrame(columns=grouped_data.columns)
    for _, group_data in grouped_data.groupby(subtotal_cols):
        key = group_data["key"].iloc[0]
        subtotal_row = subtotals[subtotals["key"] == key]
        final_data = pd.concat([final_data, subtotal_row, group_data], ignore_index=True)

    final_data.drop(columns=["key"], inplace=True)
    return final_data


def calculate_totals(grouped_data, subtotal_cols, agg_col):
    print("Processed data:")
    print(grouped_data)
    original_agg_sum = grouped_data[agg_col].sum()

    if subtotal_cols:
        final_data = prepare_subtotal_rows(grouped_data, subtotal_cols, agg_col)
    else:
        final_data = grouped_data

    grand_total_row = {agg_col: original_agg_sum,
                       subtotal_cols[0] if subtotal_cols else final_data.columns[0]: 'Grand Total'}
    grand_total = pd.DataFrame([grand_total_row], columns=final_data.columns)
    final_data = pd.concat([final_data, grand_total], ignore_index=True)

    print("subtotal column:")
    print(subtotal_cols)
    print("Aggregation column:")
    print(agg_col)
    print("Final data: ")
    print(final_data)

    return final_data
