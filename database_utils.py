
import sqlite3
import logging
import json
import pandas as pd

def get_db_connection(db_path='data.db'):
    """Get a database connection"""
    return sqlite3.connect(db_path)

def load_configuration(config_type, db_path='data.db'):
    """Load configuration data from database"""
    conn = sqlite3.connect(db_path)
    result = []
    try:
        cursor = conn.cursor()
        query = f'SELECT data FROM {config_type}'
        logging.debug(f"Executing query: {query}")
        cursor.execute(query)
        data = cursor.fetchall()
        if data:
            logging.debug(f"Data retrieved from {config_type}: {data[:3]}...")
        for item in data:
            try:
                entries = json.loads(item[0])
                for entry in entries:
                    entry['Make'] = str(entry['Make']).strip() if pd.notna(entry['Make']) else "Unknown"
                    entry['Model'] = str(entry['Model']).strip() if pd.notna(entry['Model']) else "Unknown"
                result.extend(entries)
            except json.JSONDecodeError as je:
                logging.error(f"JSON decoding error for item {item[0]}: {je}")
    except sqlite3.Error as e:
        logging.error(f"SQLite error encountered when loading configuration for {config_type}: {e}")
    finally:
        conn.close()
        return result

def get_prequal_data(db_path='data.db'):
    """Get prequal data from database"""
    return load_configuration('prequal', db_path)

def get_unique_makes(data):
    """Get unique makes from prequal data"""
    # Create a set of unique makes
    makes = set()
    seen_makes = set()  # Track makes we've already seen
    
    for item in data:
        if isinstance(item, dict) and 'Make' in item:
            make = item['Make']
            if (isinstance(make, str) and 
                make.strip() and 
                make.lower() not in ['unknown', 'nan', 'none', 'null'] and
                make.strip() not in seen_makes):  # Only add if we haven't seen it
                
                make = make.strip()
                makes.add(make)
                seen_makes.add(make)
    
    # Sort and return the list of makes
    makes_list = sorted(list(makes))
    logging.debug(f"Found {len(makes_list)} unique makes: {makes_list}")
    return makes_list

def get_unique_models(data, year, make):
    """Get unique models for a given year and make"""
    models = set()
    seen_models = set()  # Track models we've already seen
    
    for item in data:
        if isinstance(item, dict):
            try:
                if ('Year' in item and 'Make' in item and 'Model' in item and
                    str(int(float(item['Year']))) == year and
                    isinstance(item['Make'], str) and
                    item['Make'].strip() == make.strip() and
                    isinstance(item['Model'], str) and
                    item['Model'].strip() and
                    item['Model'].lower() not in ['unknown', 'nan', 'none', 'null'] and
                    item['Model'].strip() not in seen_models):  # Only add if we haven't seen it
                    
                    model = item['Model'].strip()
                    models.add(model)
                    seen_models.add(model)
            except (ValueError, TypeError):
                continue
    
    # Sort and return the list of models
    models_list = sorted(list(models))
    logging.debug(f"Found {len(models_list)} models for {year} {make}: {models_list}")
    return models_list

def get_unique_years(data):
    """Get unique years from prequal data"""
    years = set()
    seen_years = set()  # Track years we've already seen
    
    for item in data:
        if isinstance(item, dict) and 'Year' in item:
            try:
                year = int(float(item['Year']))
                if 1900 <= year <= 2100 and str(year) not in seen_years:  # Only add if we haven't seen it
                    years.add(str(year))
                    seen_years.add(str(year))
            except (ValueError, TypeError):
                continue
    
    # Sort in reverse order (newest first) and return
    years_list = sorted(list(years), reverse=True)
    logging.debug(f"Found {len(years_list)} unique years: {years_list}")
    return years_list

def populate_vehicle_dropdowns(selector, db_path='data.db'):
    """Populate the dropdowns for a vehicle selector"""
    try:
        data = get_prequal_data(db_path)
        if data:
            # Extract unique years and makes
            years = get_unique_years(data)
            makes = get_unique_makes(data)
            
            logging.debug(f"Found {len(years)} years and {len(makes)} makes")
            logging.debug(f"Years: {years}")
            logging.debug(f"Makes: {makes}")
            
            # Populate dropdowns
            selector.year.clear()
            selector.year.addItem("Select Year")
            for year in years:
                selector.year.addItem(year)
                
            selector.make.clear()
            selector.make.addItem("Select Make")
            for make in makes:
                selector.make.addItem(make)
                
    except Exception as e:
        logging.error(f"Error populating dropdowns: {e}")
        logging.exception("Full traceback:")

def update_vehicle_models(selector, year, make, db_path='data.db'):
    """Update the models dropdown based on year and make"""
    selector.model.clear()
    selector.model.addItem("Select Model")
    
    if year == "Select Year" or make == "Select Make":
        return
        
    try:
        data = get_prequal_data(db_path)
        if data:
            # Get unique models for selected year and make
            models = get_unique_models(data, year, make)
            logging.debug(f"Found {len(models)} models for {year} {make}")
            logging.debug(f"Models: {models}")
            
            for model in models:
                selector.model.addItem(model)
    except Exception as e:
        logging.error(f"Error updating models: {e}")

def get_manufacturer_chart_data(db_path='data.db'):
    """Get all manufacturer chart data from database"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='manufacturer_chart'")
        if not cursor.fetchone():
            conn.close()
            return []
        
        # Get all data
        cursor.execute("SELECT * FROM manufacturer_chart")
        columns = [description[0] for description in cursor.description]
        rows = cursor.fetchall()
        
        # Convert to list of dictionaries
        data = []
        for row in rows:
            record = dict(zip(columns, row))
            data.append(record)
        
        conn.close()
        return data
    except Exception as e:
        logging.error(f"Error getting manufacturer chart data: {e}")
        return []

def get_unique_years_from_manufacturer_chart(db_path='data.db'):
    """Get unique years from manufacturer chart data"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='manufacturer_chart'")
        if not cursor.fetchone():
            conn.close()
            return []
        
        # Get unique years
        cursor.execute("SELECT DISTINCT Year FROM manufacturer_chart WHERE Year IS NOT NULL AND Year != '' ORDER BY Year DESC")
        years = [str(row[0]) for row in cursor.fetchall()]
        
        conn.close()
        return years
    except Exception as e:
        logging.error(f"Error getting unique years from manufacturer chart: {e}")
        return []

def get_unique_makes_from_manufacturer_chart(db_path='data.db'):
    """Get unique makes from manufacturer chart data"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='manufacturer_chart'")
        if not cursor.fetchone():
            conn.close()
            return []
        
        # Get unique makes
        cursor.execute("SELECT DISTINCT Make FROM manufacturer_chart WHERE Make IS NOT NULL AND Make != '' ORDER BY Make")
        makes = [row[0] for row in cursor.fetchall()]
        
        conn.close()
        return makes
    except Exception as e:
        logging.error(f"Error getting unique makes from manufacturer chart: {e}")
        return []

def get_unique_models_from_manufacturer_chart(year, make, db_path='data.db'):
    """Get unique models from manufacturer chart data for given year and make"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='manufacturer_chart'")
        if not cursor.fetchone():
            conn.close()
            return []
        
        # Convert year to float format for database comparison (e.g., "2021" -> "2021.0")
        year_float = f"{float(year)}"
        
        # Get unique models for the given year and make
        cursor.execute("SELECT DISTINCT Model FROM manufacturer_chart WHERE Year = ? AND Make = ? AND Model IS NOT NULL AND Model != '' ORDER BY Model", (year_float, make))
        models = [row[0] for row in cursor.fetchall()]
        
        conn.close()
        return models
    except Exception as e:
        logging.error(f"Error getting unique models from manufacturer chart: {e}")
        return []

def get_vehicle_data(vehicle, db_path='data.db'):
    """Get all relevant data for a vehicle"""
    try:
        data = get_prequal_data(db_path)
        if data:
            # Filter data for selected vehicle
            prequal_data = []
            for item in data:
                if isinstance(item, dict):
                    try:
                        if ('Year' in item and 'Make' in item and 'Model' in item and
                            str(int(float(item['Year']))) == vehicle['year'] and
                            isinstance(item['Make'], str) and
                            item['Make'].strip() == vehicle['make'].strip() and
                            isinstance(item['Model'], str) and
                            item['Model'].strip() == vehicle['model'].strip()):
                            prequal_data.append(item)
                    except (ValueError, TypeError):
                        continue
            
            if prequal_data:
                return {
                    'prequal': prequal_data
                }
        
        return None
        
    except Exception as e:
        logging.error(f"Error getting vehicle data: {e}")
        return None
