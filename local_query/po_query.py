import sqlite3
poweron_db = "/Users/alan/Documents/Databases/export_of_dl12_after_scada_load_and_commissioning_and_pfl3.db"

def getComponentIdFromAlias(alias):
    if alias is None or alias == "":
        return ""
    # handle alias is nan
    if alias != alias:
        return ""
    
    db = sqlite3.connect(poweron_db)
    cursor = db.cursor()
    
    try:
        query = """
                SELECT  component_id COMPONENT_ID 
                FROM    component_header 
                WHERE   component_alias = ? 
            """
        cursor.execute(query, (alias, ))
        row = cursor.fetchone()
        result = row[0] if row else ""
        
    except sqlite3.Error as e:
        print(f"Error in getComponentIdFromAlias for alias = '{alias}', with query '{query}': {e}")
        result = ""
        
    finally:
        cursor.close()
        db.close()
        
    return result


def check_if_component_alias_exists_in_poweron(component_alias: str) -> bool:
    """
    Check if a component alias exists in the PowerOn database.
    """

    try:
        id = getComponentIdFromAlias(component_alias)
        exists = 1 if id is not None and id != "" else 0
    except:
        print(f"Error in check_if_component_alias_exists_in_poweron for alias = '{component_alias}'")
        return False

    return exists


def checkIfComponentAliasInScanPointComponents(component_alias):
    if component_alias is None or component_alias == "":
        return False
    # handle alias is nan
    if component_alias != component_alias:
        return False
    
    db = sqlite3.connect(poweron_db)
    cursor = db.cursor()
    
    try:
        query = "SELECT 1 FROM scan_point_components WHERE COMPONENT_ALIAS = ?"
        cursor.execute(query, (component_alias, ))
        row = cursor.fetchone()
        return True if row else False
    except sqlite3.Error as e:
        print(f"Error in checkIfComponentAliasInScanPointComponents for alias = '{component_alias}', with query 'SELECT 1 FROM scan_point_components WHERE COMPONENT_ALIAS = ?': {e}")
        return False