import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

def load_and_filter_data(file_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Load the Excel file and filter for European data.
    
    Args:
        file_path (str): Path to the Excel file
        sheet_name (str): Name of the sheet to read
        
    Returns:
        pd.DataFrame: Filtered dataframe containing only European data
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Filter for Europe
        europe_df = df[df['Continent'] == 'Europe'].copy()
        
        # Handle missing values in Trade Value
        europe_df['Trade Value'] = pd.to_numeric(europe_df['Trade Value'], errors='coerce')
        europe_df = europe_df.dropna(subset=['Trade Value'])
        
        return europe_df
    
    except FileNotFoundError:
        print(f"Error: File {file_path} not found.")
        return None
    except Exception as e:
        print(f"Error loading data: {str(e)}")
        return None

def calculate_statistics(df: pd.DataFrame) -> dict:
    """
    Calculate key statistics from the trade value data.
    
    Args:
        df (pd.DataFrame): Input dataframe with Trade Value column
        
    Returns:
        dict: Dictionary containing calculated statistics
    """
    stats = {
        'mean': df['Trade Value'].mean(),
        'std': df['Trade Value'].std(),
        'min': df['Trade Value'].min(),
        'max': df['Trade Value'].max(),
        'range': df['Trade Value'].max() - df['Trade Value'].min()
    }
    return stats

def find_countries_in_range(df: pd.DataFrame, percentage: float = 10) -> pd.DataFrame:
    """
    Find countries with export values within specified percentage of max value.
    
    Args:
        df (pd.DataFrame): Input dataframe
        percentage (float): Percentage threshold (default: 10)
        
    Returns:
        pd.DataFrame: Filtered dataframe with countries in range
    """
    max_value = df['Trade Value'].max()
    threshold = max_value * (1 - percentage/100)
    
    return df[df['Trade Value'] >= threshold].sort_values('Trade Value', ascending=False)

def create_visualizations(df: pd.DataFrame, output_dir: str = 'output'):
    """
    Create and save visualizations of the data.
    
    Args:
        df (pd.DataFrame): Input dataframe
        output_dir (str): Directory to save visualizations
    """
    # Create output directory if it doesn't exist
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # Set style to a default matplotlib style
    plt.style.use('default')
    
    # Format trade values to billions for better readability
    df['Trade Value (Billions)'] = df['Trade Value'] / 1e9
    
    # Create bar plot of trade values by country
    plt.figure(figsize=(12, 6))
    sns.barplot(data=df, x='Country', y='Trade Value (Billions)')
    plt.xticks(rotation=45, ha='right')
    plt.title('Crude Petroleum Export Values by European Country')
    plt.ylabel('Trade Value (Billions USD)')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f'{output_dir}/export_values_by_country.png', dpi=300, bbox_inches='tight')
    plt.close()
    
    # Create histogram of trade values
    plt.figure(figsize=(10, 6))
    sns.histplot(data=df, x='Trade Value (Billions)', bins=15)
    plt.title('Distribution of Crude Petroleum Export Values in Europe')
    plt.xlabel('Trade Value (Billions USD)')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(f'{output_dir}/export_values_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

def main():
    # Configuration
    file_path = 'data.xlsx'  # Update this with your file path
    sheet_name = 'Exporters-of-Crude-Petroleum-2'
    
    # Load and filter data
    df = load_and_filter_data(file_path, sheet_name)
    if df is None:
        return
    
    # Calculate statistics
    stats = calculate_statistics(df)
    print("\nSummary Statistics (in USD):")
    print("-" * 40)
    for key, value in stats.items():
        # Convert to billions for better readability
        value_in_billions = value / 1e9
        print(f"{key.capitalize():15}: {value_in_billions:,.2f} billion")
    
    # Find countries within 10% of max value
    top_exporters = find_countries_in_range(df, percentage=10)
    print("\nTop Exporters (within 10% of maximum value):")
    print("-" * 60)
    print(f"{'Country':15} {'ISO 3':8} {'Trade Value (Billions USD)':>25}")
    print("-" * 60)
    for _, row in top_exporters.iterrows():
        trade_value_billions = row['Trade Value'] / 1e9
        print(f"{row['Country']:15} {row['ISO 3']:8} {trade_value_billions:>25,.2f}")
    
    # Create visualizations
    create_visualizations(df)
    print("\nVisualizations have been saved to the 'output' directory.")

if __name__ == "__main__":
    main()
