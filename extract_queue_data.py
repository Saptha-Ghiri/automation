import pandas as pd
from collections import defaultdict
import json
import os

def normalize_status(status):
    """
    Normalize similar statuses to group them together
    """
    if not status:
        return status

    status_lower = str(status).lower().strip()

    # Group approval-related statuses
    if any(word in status_lower for word in ['approval', 'approve', 'awaiting approval', 'pending approval']):
        return 'Approval'

    # Group closed-related statuses
    if any(word in status_lower for word in ['closed', 'ticket closed', 'close']):
        return 'Closed'

    # Group resolved-related statuses
    if any(word in status_lower for word in ['resolved', 'resolve', 'completed', 'complete']):
        return 'Resolved'

    # Group in-progress related statuses
    if any(word in status_lower for word in ['inprogress', 'in progress', 'in-progress', 'progress', 'working']):
        return 'In Progress'

    # Group new/open related statuses
    if any(word in status_lower for word in ['new', 'open', 'created']):
        return 'New'

    # Group awaiting/waiting related statuses (excluding approval which is handled above)
    if any(word in status_lower for word in ['awaiting', 'waiting', 'pending']) and 'approval' not in status_lower:
        return 'Awaiting'

    # Return original status if no grouping applies
    return status

def extract_resource_status_counts(file_path):
    """
    Extract resource, status and date-wise count data from Excel file
    Returns dictionaries with counts for each resource, status, and date-wise breakdown
    """
    try:
        # Try to read the Excel file
        df = pd.read_excel(file_path)
        print(f"File read successfully! Shape: {df.shape}")
        
    except PermissionError:
        print(f"Permission denied for {file_path}")
        print("The file might be open in Excel. Please close it and try again.")
        return None, None, None
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None, None, None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None, None, None
    
    print("\nColumn names in the file:")
    for i, col in enumerate(df.columns):
        print(f"{i:2d}: {col}")
    
    print("\nFirst few rows:")
    print(df.head().to_string())
    
    # Initialize dictionaries to store counts
    resource_counts = defaultdict(int)
    status_counts = defaultdict(int)
    date_wise_data = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    
    # Check for possible column names (case insensitive)
    resource_col = None
    status_col = None
    date_col = None
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        
        # Check for resource/user column
        if any(word in col_lower for word in ['resource', 'assigned', 'user', 'owner', 'responsible', 'assignee']):
            resource_col = col
            print(f"\nFound resource column: '{col}'")
        
        # Check for status column  
        if any(word in col_lower for word in ['status', 'state', 'condition']):
            status_col = col
            print(f"Found status column: '{col}'")
        
        # Check for date column
        if any(word in col_lower for word in ['date', 'created', 'updated', 'modified', 'timestamp']):
            date_col = col
            print(f"Found date column: '{col}'")
    
    # If columns not found automatically, use manual assignment
    if not resource_col or not status_col:
        print("\nCould not automatically identify columns.")
        print("Please manually specify column names or indices in the script.")
        
        # Try common patterns
        if len(df.columns) >= 2:
            if not resource_col:
                resource_col = df.columns[0]  # Assume first column is resource
                print(f"Using first column as resource: '{resource_col}'")
            
            if not status_col:
                # Look for a column with limited unique values (likely status)
                for col in df.columns[1:]:
                    if len(df[col].dropna().unique()) < 15:  # Status usually has few unique values
                        status_col = col
                        print(f"Using column as status: '{status_col}'")
                        break
            
            if not date_col:
                # Look for date-like columns
                for col in df.columns:
                    if df[col].dtype == 'datetime64[ns]' or 'date' in str(col).lower():
                        date_col = col
                        print(f"Using column as date: '{date_col}'")
                        break
    
    # Process the data - Handle merged date cells
    current_date = None
    
    for index, row in df.iterrows():
        # Get resource
        resource = None
        if resource_col and pd.notna(row.get(resource_col)):
            resource = str(row[resource_col]).strip()
            if resource:
                resource_counts[resource] += 1
        
        # Get status
        status = None
        if status_col and pd.notna(row.get(status_col)):
            original_status = str(row[status_col]).strip()
            if original_status:
                status = normalize_status(original_status)
                status_counts[status] += 1
        
        # Get date - Handle merged cells
        date_str = None
        if date_col and pd.notna(row.get(date_col)):
            date_val = row[date_col]
            if pd.api.types.is_datetime64_any_dtype(type(date_val)):
                date_str = date_val.strftime('%d/%m/%Y')
            else:
                # Try to parse as string date
                try:
                    parsed_date = pd.to_datetime(str(date_val))
                    date_str = parsed_date.strftime('%d/%m/%Y')
                except:
                    date_str = str(date_val)
            
            # Update current_date when we find a non-empty date
            if date_str and date_str.strip():
                current_date = date_str
        
        # Use current_date for merged cells (when date cell is empty but we have a current date)
        if not date_str and current_date:
            date_str = current_date
        
        # Build date-wise data using the resolved date
        if date_str and resource:
            date_wise_data[date_str]['resources'][resource] += 1
        if date_str and status:
            date_wise_data[date_str]['statuses'][status] += 1
    
    # Convert to regular dicts
    resource_counts = dict(resource_counts)
    status_counts = dict(status_counts)
    
    # Convert date_wise_data to regular nested dicts
    date_wise_regular = {}
    for date, data in date_wise_data.items():
        date_wise_regular[date] = {
            'resources': dict(data['resources']),
            'statuses': dict(data['statuses'])
        }
    
    return resource_counts, status_counts, date_wise_regular

def create_sample_data():
    """Create sample data if file cannot be read"""
    resource_counts = {
        "Abhijeet Nashikkar": 25,
        "Aditya Anand": 18, 
        "Nishanth Senthilkumar": 22,
        "Sakthivel s Venkatachalam": 15,
        "Saptha": 8
    }
    
    status_counts = {
        "New": 12,
        "In Progress": 20,
        "Awaiting": 15,
        "Internal Solution Provided": 10,
        "Resolved": 25,
        "Closed": 6
    }
    
    # Sample date-wise data
    date_wise_data = {
        "09/01/2025": {
            "resources": {
                "Abhijeet Nashikkar": 5,
                "Aditya Anand": 4,
                "Nishanth Senthilkumar": 6,
                "Sakthivel s Venkatachalam": 3,
                "Saptha": 2
            },
            "statuses": {
                "New": 3,
                "In Progress": 5,
                "Awaiting": 4,
                "Internal Solution Provided": 2,
                "Resolved": 5,
                "Closed": 1
            }
        },
        "09/02/2025": {
            "resources": {
                "Abhijeet Nashikkar": 4,
                "Aditya Anand": 3,
                "Nishanth Senthilkumar": 4,
                "Sakthivel s Venkatachalam": 2,
                "Saptha": 1
            },
            "statuses": {
                "New": 2,
                "In Progress": 3,
                "Awaiting": 2,
                "Internal Solution Provided": 2,
                "Resolved": 4,
                "Closed": 1
            }
        },
        "09/03/2025": {
            "resources": {
                "Abhijeet Nashikkar": 6,
                "Aditya Anand": 4,
                "Nishanth Senthilkumar": 5,
                "Sakthivel s Venkatachalam": 4,
                "Saptha": 2
            },
            "statuses": {
                "New": 2,
                "In Progress": 4,
                "Awaiting": 3,
                "Internal Solution Provided": 2,
                "Resolved": 8,
                "Closed": 2
            }
        },
        "09/04/2025": {
            "resources": {
                "Abhijeet Nashikkar": 5,
                "Aditya Anand": 3,
                "Nishanth Senthilkumar": 4,
                "Sakthivel s Venkatachalam": 3,
                "Saptha": 1
            },
            "statuses": {
                "New": 2,
                "In Progress": 4,
                "Awaiting": 3,
                "Internal Solution Provided": 2,
                "Resolved": 4,
                "Closed": 1
            }
        },
        "09/05/2025": {
            "resources": {
                "Abhijeet Nashikkar": 5,
                "Aditya Anand": 4,
                "Nishanth Senthilkumar": 3,
                "Sakthivel s Venkatachalam": 3,
                "Saptha": 2
            },
            "statuses": {
                "New": 3,
                "In Progress": 4,
                "Awaiting": 3,
                "Internal Solution Provided": 2,
                "Resolved": 4,
                "Closed": 1
            }
        }
    }
    
    return resource_counts, status_counts, date_wise_data

def main():
    file_path = r"C:\Users\sapth1504421\OneDrive - Mastek Limited\Desktop\devops_projects\Weekly_report_automation\inputs\temp_daas_queue.xlsx"
    
    print("Extracting data from temp_daas_queue.xlsx...")
    print("=" * 50)
    
    # Try to extract from file
    resource_counts, status_counts, date_wise_data = extract_resource_status_counts(file_path)
    print("\nExtraction complete.")
    print("resource_counts:", resource_counts)
    print("status_counts:", status_counts)
    print("date_wise_data:", date_wise_data)    
    print()
    # If extraction failed, use sample data
    if resource_counts is None or status_counts is None or date_wise_data is None:
        print("\nUsing sample data...")
        resource_counts, status_counts, date_wise_data = create_sample_data()
    
    # Display results
    print("\nRESOURCE COUNTS:")
    print("-" * 30)
    for resource, count in sorted(resource_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"{resource:25} : {count}")
    
    print("\nSTATUS COUNTS:")
    print("-" * 30)
    for status, count in sorted(status_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"{status:25} : {count}")
    
    # Display date-wise data
    print("\nDATE-WISE BREAKDOWN:")
    print("=" * 50)
    for date in sorted(date_wise_data.keys()):
        print(f"\nDate: {date}")
        print("-" * 20)
        
        print("  Resources:")
        for resource, count in sorted(date_wise_data[date]['resources'].items(), key=lambda x: x[1], reverse=True):
            print(f"    {resource:20} : {count}")
        
        print("  Statuses:")
        for status, count in sorted(date_wise_data[date]['statuses'].items(), key=lambda x: x[1], reverse=True):
            print(f"    {status:20} : {count}")
    
    # Summary
    total_resources = sum(resource_counts.values())
    total_statuses = sum(status_counts.values())
    total_dates = len(date_wise_data)
    
    print(f"\nSUMMARY:")
    print(f"Total records (by resource): {total_resources}")
    print(f"Total records (by status): {total_statuses}")
    print(f"Unique resources: {len(resource_counts)}")
    print(f"Unique statuses: {len(status_counts)}")
    print(f"Unique dates: {total_dates}")
    
    # Save to JSON
    output_data = {
        "resource_counts": resource_counts,
        "status_counts": status_counts,
        "date_wise_data": date_wise_data,
        "summary": {
            "total_by_resource": total_resources,
            "total_by_status": total_statuses,
            "unique_resources": len(resource_counts),
            "unique_statuses": len(status_counts),
            "unique_dates": total_dates
        }
    }
    
    try:
        with open("daas_queue_data.json", 'w') as f:
            json.dump(output_data, f, indent=4)
        print("\nData saved to: daas_queue_data.json")
    except Exception as e:
        print(f"Error saving JSON: {e}")
    
    # Print Python dictionaries
    print("\n" + "="*50)
    print("PYTHON DICTIONARIES:")
    print("="*50)
    
    print("\nresource_counts = {")
    for resource, count in resource_counts.items():
        print(f'    "{resource}": {count},')
    print("}")
    
    print("\nstatus_counts = {")
    for status, count in status_counts.items():
        print(f'    "{status}": {count},')
    print("}")
    
    print("\n# Date-wise data")
    print("date_wise_data = {")
    for date, data in date_wise_data.items():
        print(f'    "{date}": {{')
        print(f'        "resources": {{')
        for resource, count in data['resources'].items():
            print(f'            "{resource}": {count},')
        print(f'        }},')
        print(f'        "statuses": {{')
        for status, count in data['statuses'].items():
            print(f'            "{status}": {count},')
        print(f'        }}')
        print(f'    }},')
    print("}")
    
    return resource_counts, status_counts, date_wise_data

if __name__ == "__main__":
    resource_counts, status_counts, date_wise_data = main()
    print("\nScript completed!")