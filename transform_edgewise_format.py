"""
Transform VideoProcessingOutput/EdgeWise_DirectionalCounts_5MinIntervalColumns_Datewise.xlsx
from wide format (interval columns) to long format matching Streamlit_Callibration/5minCompileSummary_UPDATED.xlsx
"""
import pandas as pd
from pathlib import Path

def extract_time_from_interval(interval: str) -> tuple[str, str]:
    """Extract TimeStart and TimeEnd from interval like '08:00:00-08:05:00'"""
    parts = interval.split('-')
    if len(parts) == 2:
        return parts[0], parts[1]
    return None, None

def transform_edgewise_to_summary():
    src_file = Path('VideoProcessingOutput/EdgeWise_DirectionalCounts_5MinIntervalColumns_Datewise.xlsx')
    output_file = Path('VideoProcessingOutput/EdgeWise_DirectionalCounts_Summary_Format.xlsx')
    
    # Read all four sheets
    sheets_mapping = {
        'Bus_5MinCols': 'Bus',
        'Car_5MinCols': 'Car',
        'Motorcycle_5MinCols': 'Motorcycle',
        'Total_Volume_5MinCols': 'Total',
    }
    
    all_rows = []
    
    for sheet_name, vehicle_class in sheets_mapping.items():
        df = pd.read_excel(src_file, sheet_name=sheet_name)
        
        # Identify interval columns (those with format HH:MM:SS-HH:MM:SS)
        key_cols = ['Location', 'Date', 'Approach', 'Direction', 'SourceWorkbook']
        interval_cols = [col for col in df.columns if col not in key_cols]
        
        # Melt the dataframe from wide to long
        melted = pd.melt(
            df,
            id_vars=key_cols,
            value_vars=interval_cols,
            var_name='Interval',
            value_name='Count'
        )
        
        # Extract TimeStart and TimeEnd from interval
        melted[['TimeStart', 'TimeEnd']] = melted['Interval'].apply(
            lambda x: pd.Series(extract_time_from_interval(x))
        )
        
        # Create SessionFolder from Date (interpret as folder like 20260222)
        melted['SessionFolder'] = melted['Date'].str.replace('-', '')
        
        # Create IntervalFolder from TimeStart (hour part, e.g., 08 from 08:00:00)
        melted['IntervalFolder'] = melted['TimeStart'].str.split(':').str[0]
        
        # Add VehicleClass (only for non-Total sheets)
        if vehicle_class != 'Total':
            melted['VehicleClass'] = vehicle_class
        else:
            melted['VehicleClass'] = 'Total'
        
        # Select and reorder columns to match target format
        melted = melted[[
            'Location', 'SessionFolder', 'IntervalFolder',
            'TimeStart', 'TimeEnd',
            'Approach', 'Direction',
            'VehicleClass', 'Count', 'SourceWorkbook'
        ]]
        
        all_rows.append(melted)
    
    # Combine all sheets
    summary_df = pd.concat(all_rows, ignore_index=True)
    
    # Convert Count to numeric
    summary_df['Count'] = pd.to_numeric(summary_df['Count'], errors='coerce').fillna(0)
    
    # Write to Excel with single Summary sheet
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f'Output file: {output_file}')
    print(f'Total rows: {len(summary_df)}')
    print(f'Columns: {list(summary_df.columns)}')
    print(f'Unique VehicleClasses: {summary_df["VehicleClass"].unique()}')
    print(f'Sample rows:\n{summary_df.head(3)}')

if __name__ == '__main__':
    transform_edgewise_to_summary()
