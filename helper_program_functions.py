import pandas as pd


def filter_out_items_booked_to_0004_spec_cust_requirement_location(mb52_df, vl10d_merged_df):
    mb52_df_copy = mb52_df.copy()
    # drop rows with NaN in SAP_nr column
    mb52_df_copy.dropna(subset=["Numer zapasu specjalnego"], inplace=True)
    mb52_df_copy = mb52_df_copy[mb52_df_copy['storage_loc'] == "0004"]

    # drop rows which were already booked to 0004 or were produced to 0004
    # Rename 'stock' to 'quantity' in mb52_df_copy temporarily for comparison
    mb52_df_copy.rename(columns={'stock': 'quantity'}, inplace=True)

    # set appropriate data types
    vl10d_merged_df['quantity'] = pd.to_numeric(vl10d_merged_df['quantity'], errors='coerce').round().astype('Int64')
    mb52_df_copy['quantity'] = pd.to_numeric(mb52_df_copy['quantity'], errors='coerce').round().astype('Int64')

    # Perform an inner merge to find matching rows
    matching_rows = pd.merge(
        vl10d_merged_df.reset_index(),
        mb52_df_copy,
        on=['SAP_nr', 'quantity', 'document_number'],
        how='inner'
    )

    # matching_rows
    # Drop the matching rows from vl10d_merged_df
    vl10d_merged_df = vl10d_merged_df.drop(matching_rows['index'])

    return vl10d_merged_df


def fill_storage_location_quantities(mb52_df, vl10d_merged_df):
    # Keep only rows where 'Numer zapasu specjalnego' has NaN values
    mb52_df = mb52_df[mb52_df['Numer zapasu specjalnego'].isnull()]

    for row in mb52_df.iterrows():
        stock = row[1]['stock']
        sap_nr = row[1]['SAP_nr']
        storage_loc = row[1]['storage_loc']
        vl10d_merged_df.loc[vl10d_merged_df['SAP_nr'] == sap_nr, f'loc_{storage_loc}'] = stock

    return vl10d_merged_df
