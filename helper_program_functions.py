import pandas as pd
from gui_manager import show_message


def filter_out_items_booked_to_0004_spec_cust_requirement_location(mb52_df, vl10x_merged_df):
    mb52_df_copy = mb52_df.copy()
    # drop rows with NaN in SAP_nr column
    mb52_df_copy.dropna(subset=["Numer zapasu specjalnego"], inplace=True)
    mb52_df_copy = mb52_df_copy[mb52_df_copy['storage_loc'] == "0004"]

    # get doc_num and doc_pos
    mb52_df_copy['document_number'] = mb52_df_copy['Numer zapasu specjalnego'].apply(lambda x: x.split('/')[0])
    mb52_df_copy['doc_position'] = mb52_df_copy['Numer zapasu specjalnego'].apply(lambda x: x.split('/')[1].strip())

    # drop rows which were already booked to 0004 or were produced to 0004
    # Rename 'stock' to 'quantity' in mb52_df_copy temporarily for comparison
    mb52_df_copy.rename(columns={'stock': 'quantity'}, inplace=True)

    # set appropriate data types
    mb52_df_copy['quantity'] = mb52_df_copy['quantity'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.').strip()))
    vl10x_merged_df['quantity'] = pd.to_numeric(vl10x_merged_df['quantity'], errors='coerce').astype('float')
    mb52_df_copy['quantity'] = pd.to_numeric(mb52_df_copy['quantity'], errors='coerce').astype('float')

    # Perform an inner merge to find matching rows
    matching_rows = pd.merge(
        vl10x_merged_df.reset_index(),
        mb52_df_copy,
        on=['SAP_nr', 'quantity', 'document_number'],
        how='inner'
    )

    # matching_rows
    # Drop the matching rows from vl10d_merged_df
    vl10x_merged_df = vl10x_merged_df.drop(matching_rows['index'])

    return vl10x_merged_df


def fill_storage_location_quantities(mb52_df, vl10x_merged_df):
    # Keep only rows where 'Numer zapasu specjalnego' has NaN values
    mb52_df = mb52_df[mb52_df['Numer zapasu specjalnego'].isnull()]

    for row in mb52_df.iterrows():
        stock = str(row[1]['stock']).replace('.', ',')
        sap_nr = row[1]['SAP_nr']
        storage_loc = row[1]['storage_loc']
        vl10x_merged_df.loc[vl10x_merged_df['SAP_nr'] == sap_nr, f'loc_{storage_loc}'] = stock

    return vl10x_merged_df


def get_source_storage_location(row, quantity):
    storage_locs = ['loc_0007', 'loc_0003', 'loc_0750', 'loc_0005']
    for loc in storage_locs:
        if float(str(row[loc]).replace(',', '.')) >= float(str(quantity).strip()):
            return loc[-4:]  # Return the last 4 characters of the location name
    return None


def determine_header_suffix(row):
    prod_suffix = 'produkcja'
    service_suffix = 'serwis'

    production_strings = ('TLBTL')
    service_strings = ('GSB', 'FO')
    quantity_threshold = 30

    if str(row['product_name']).startswith(production_strings):
        return prod_suffix
    elif str(row['product_name']).startswith(service_strings):
        return service_suffix
    else:
        if row['quantity'] > quantity_threshold:
            return prod_suffix
        else:
            return service_suffix


def determine_vl10c_header(row, sales_offices_map):
    if row['SAP_nr'] == '773630':
        return 'BelatronicUPS'
    else:
        return sales_offices_map[row['sales_office']]
