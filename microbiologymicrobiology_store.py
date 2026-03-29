import os
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')

# =====================================================
# Configuration
# =====================================================
INVENTORY_FILE = "microbiology_inventory.xlsx"
TRANSACTIONS_FILE = "microbiology_transactions.xlsx"
MEDIA_PREP_FILE = "media_preparation.xlsx"
REPORTS_DIR = "microbiology_reports"

# Create directories if they don't exist
os.makedirs(REPORTS_DIR, exist_ok=True)

# Main categories
MAIN_CATEGORIES = [
    "Strains (Reference Materials)",
    "Disposal (Petri dishes, tissues, membrane filters, etc)",
    "Culture Media (Broths, Agars, etc)",
    "Chemicals & Reagents",
    "Equipment & Instruments",
    "Other"
]

# =====================================================
# Category-specific columns
# =====================================================

# Culture Media columns
CULTURE_MEDIA_COLUMNS = [
    "Item Code", "Item Name", "Category", "Type of Media", "Lot No.",
    "Reference No.", "Date of Opening", "Expiry Date", "Storage Conditions",
    "Quantity (g)", "For 1000mL (g)", "Current Quantity (g)", "Unit", "Notes"
]

# Strains columns
STRAINS_COLUMNS = [
    "Item Code", "Item Name", "Category", "Reference Material Name", "Lot No.",
    "Manufacturer", "Recommended Storage Conditions", "Validity Date",
    "Quantity", "Unit", "Passage Number", "Date Received", "Notes"
]

# Disposal columns (consumables)
DISPOSAL_COLUMNS = [
    "Item Code", "Item Name", "Category", "Description", "Quantity",
    "Unit", "Package Size", "Manufacturer", "Catalog No.", "Storage Location",
    "Min Stock Level", "Date Received", "Notes"
]

# Chemicals columns
CHEMICALS_COLUMNS = [
    "Item Code", "Item Name", "Category", "CAS No.", "Lot No.",
    "Manufacturer", "Grade", "Quantity", "Unit", "Storage Conditions",
    "Hazard Symbols", "MSDS Available", "Expiry Date", "Notes"
]

# Equipment columns
EQUIPMENT_COLUMNS = [
    "Item Code", "Item Name", "Category", "Model No.", "Serial No.",
    "Manufacturer", "Location", "Calibration Due", "Maintenance Schedule",
    "Status", "Purchase Date", "Notes"
]

# =====================================================
# Units and other lists
# =====================================================
MEDIA_TYPES = [
    "Nutrient Agar", "MacConkey Agar", "Blood Agar", "Chocolate Agar",
    "Sabouraud Dextrose Agar", "Muller Hinton Agar", "TCBS Agar",
    "XLD Agar", "Cetrimide Agar", "Luria-Bertani Broth", "Tryptic Soy Broth",
    "Brain Heart Infusion Broth", "Peptone Water", "Buffered Peptone Water",
    "R2A Agar", "Reasoner's 2A Agar", "Other"
]

STORAGE_CONDITIONS = [
    "Room Temperature (15-25°C)",
    "Refrigerated (2-8°C)",
    "Frozen (-20°C)",
    "Deep Frozen (-80°C)",
    "Desiccated",
    "Light Sensitive",
    "Other"
]

UNITS = [
    "g", "kg", "mL", "L", "each", "pack", "box", "vial", "strip", "Other"
]

DISPOSAL_UNITS = [
    "pack (100)", "pack (500)", "pack (1000)", "box", "case", "each", "roll", "Other"
]

# Master list of all possible columns (for initialization)
ALL_COLUMNS = list(set(
    ["Item Code", "Item Name", "Category"] +
    CULTURE_MEDIA_COLUMNS +
    STRAINS_COLUMNS +
    DISPOSAL_COLUMNS +
    CHEMICALS_COLUMNS +
    EQUIPMENT_COLUMNS
))

# =====================================================
# Helper Functions
# =====================================================

def load_inventory():
    """Load inventory data from Excel file"""
    if os.path.exists(INVENTORY_FILE):
        df = pd.read_excel(INVENTORY_FILE)
        # Ensure all columns exist
        for col in ALL_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df
    else:
        # Create empty DataFrame with all possible columns
        return pd.DataFrame(columns=ALL_COLUMNS)

def save_inventory(df):
    """Save inventory data to Excel file"""
    df.to_excel(INVENTORY_FILE, index=False)

def load_transactions():
    """Load transaction history"""
    if os.path.exists(TRANSACTIONS_FILE):
        df = pd.read_excel(TRANSACTIONS_FILE)
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        return df
    else:
        return pd.DataFrame(columns=["Transaction ID", "Date", "Type", "Item Code", "Item Name", "Quantity", "Unit", "Purpose", "User"])

def save_transactions(df):
    """Save transaction history"""
    df.to_excel(TRANSACTIONS_FILE, index=False)

def load_media_prep():
    """Load media preparation records"""
    if os.path.exists(MEDIA_PREP_FILE):
        return pd.read_excel(MEDIA_PREP_FILE)
    else:
        return pd.DataFrame(columns=["Date", "Media Type", "Lot No.", "Quantity (mL)", "Media Used (g)", "Water Used (mL)", "Prepared By", "Expiry Date"])

def save_media_prep(df):
    """Save media preparation records"""
    df.to_excel(MEDIA_PREP_FILE, index=False)

def generate_item_code(df, category_prefix):
    """Generate a new item code with category prefix"""
    prefix_map = {
        "Strains (Reference Materials)": "STR",
        "Disposal (Petri dishes, tissues, membrane filters, etc)": "DSP",
        "Culture Media (Broths, Agars, etc)": "MED",
        "Chemicals & Reagents": "CHM",
        "Equipment & Instruments": "EQP",
        "Other": "OTH"
    }
    
    prefix = prefix_map.get(category_prefix, "ITM")
    
    if df.empty or df["Item Code"].isna().all():
        return f"{prefix}-0001"
    else:
        # Filter items with same prefix
        df_valid = df[df["Item Code"].notna()]
        same_prefix = df_valid[df_valid["Item Code"].str.startswith(prefix, na=False)]
        if same_prefix.empty:
            return f"{prefix}-0001"
        
        last_code = same_prefix["Item Code"].iloc[-1]
        try:
            num = int(last_code.split("-")[1])
            new_num = num + 1
            return f"{prefix}-{new_num:04d}"
        except:
            return f"{prefix}-{len(same_prefix)+1:04d}"

def add_transaction(trans_type, item_code, item_name, quantity, unit, purpose=""):
    """Add a new transaction"""
    df_trans = load_transactions()
    trans_id = f"TR-{len(df_trans) + 1:06d}"
    
    new_trans = pd.DataFrame([[
        trans_id, date.today(), trans_type, item_code,
        item_name, quantity, unit, purpose, "admin"
    ]], columns=["Transaction ID", "Date", "Type", "Item Code", "Item Name", "Quantity", "Unit", "Purpose", "User"])
    
    df_trans = pd.concat([df_trans, new_trans], ignore_index=True)
    save_transactions(df_trans)

def calculate_media_needed(volume_ml, per_1000ml_g):
    """Calculate grams of media needed for a specific volume"""
    if pd.isna(per_1000ml_g) or per_1000ml_g <= 0:
        return 0
    return (volume_ml / 1000) * per_1000ml_g

def create_new_item_row(item_data, columns):
    """Create a new DataFrame row with all columns"""
    new_row = {}
    # Initialize all columns with empty string
    for col in ALL_COLUMNS:
        new_row[col] = ""
    
    # Fill in the provided data
    for i, col in enumerate(columns):
        if i < len(item_data) and col in new_row:
            new_row[col] = item_data[i]
    
    return pd.DataFrame([new_row])

# =====================================================
# Streamlit App Configuration
# =====================================================
st.set_page_config(
    page_title="Microbiology Lab Store",
    page_icon="🧫",
    layout="wide"
)

# Header
st.title("🧫 Microbiology Laboratory Store Management System")
st.markdown("---")

# =====================================================
# Sidebar Navigation
# =====================================================
st.sidebar.title("📋 Navigation")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Main Menu",
    ["Dashboard", "Add Material", "Prepare Media", "Issue Material", "Inventory", "Transactions", "Reports"]
)

st.sidebar.markdown("---")

# =====================================================
# Add Material Section
# =====================================================
if menu == "Add Material":
    st.header("➕ Add New Material to Inventory")
    
    # First: Select Category
    category = st.selectbox(
        "Select Material Category *",
        MAIN_CATEGORIES,
        key="main_category"
    )
    
    st.markdown("---")
    
    # Load existing inventory
    df_inv = load_inventory()
    
    # =========================================
    # CULTURE MEDIA
    # =========================================
    if category == "Culture Media (Broths, Agars, etc)":
        st.subheader("📊 Culture Media Entry")
        
        with st.form("culture_media_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                item_name = st.text_input("Media Name *", placeholder="e.g., Nutrient Agar")
                media_type = st.selectbox("Type of Media *", MEDIA_TYPES)
                if media_type == "Other":
                    media_type = st.text_input("Specify Media Type")
                
                lot_no = st.text_input("Lot/Batch No. *", placeholder="e.g., 12345-AB")
                ref_no = st.text_input("Reference/Catalog No.", placeholder="e.g., CM0003")
                
            with col2:
                date_opened = st.date_input("Date of Opening", value=date.today())
                expiry_date = st.date_input("Expiry Date *")
                storage_conditions = st.selectbox("Storage Conditions *", STORAGE_CONDITIONS)
            
            st.markdown("### 📦 Inventory Information")
            col3, col4 = st.columns(2)
            
            with col3:
                total_quantity_g = st.number_input(
                    "Total Quantity Received (grams) *",
                    min_value=0.0,
                    step=1.0,
                    format="%.2f",
                    help="Total amount of media received in grams"
                )
            
            with col4:
                unit = st.selectbox("Unit", ["g", "kg"], index=0)
            
            st.markdown("### 🧪 Media Preparation Calculation Reference")
            st.info("This field is used for calculating how much media is needed when preparing solutions")
            
            col5, col6 = st.columns(2)
            with col5:
                per_1000ml_g = st.number_input(
                    "For 1000 mL (grams) *",
                    min_value=0.0,
                    step=0.1,
                    format="%.2f",
                    help="How many grams of this media are required to prepare 1000 mL of solution"
                )
                
                # Show example calculation
                if per_1000ml_g > 0:
                    example_volumes = [250, 500, 1000, 2000]
                    st.caption("Example calculations:")
                    for vol in example_volumes:
                        needed = (vol / 1000) * per_1000ml_g
                        st.caption(f"  • {vol} mL → {needed:.2f} g")
            
            with col6:
                notes = st.text_area("Additional Notes", placeholder="Any special instructions or observations")
            
            st.markdown("---")
            submitted = st.form_submit_button("💾 Save Culture Media", use_container_width=True)
            
            if submitted:
                # Validation
                errors = []
                if not item_name:
                    errors.append("❌ Media Name is required")
                if not lot_no:
                    errors.append("❌ Lot/Batch No. is required")
                if not expiry_date:
                    errors.append("❌ Expiry Date is required")
                if total_quantity_g <= 0:
                    errors.append("❌ Total Quantity Received must be greater than zero")
                if per_1000ml_g <= 0:
                    errors.append("❌ For 1000mL value must be greater than zero")
                
                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    # Generate item code
                    item_code = generate_item_code(df_inv, category)
                    
                    # Prepare data for new item
                    new_item_data = [
                        item_code, item_name, category, media_type, lot_no,
                        ref_no, date_opened, expiry_date, storage_conditions,
                        total_quantity_g,  # Total quantity received
                        per_1000ml_g,      # For 1000mL calculation reference
                        total_quantity_g,  # Current quantity (starts as total received)
                        unit, 
                        notes
                    ]
                    
                    # Create new row with all columns
                    new_item = create_new_item_row(new_item_data, CULTURE_MEDIA_COLUMNS)
                    
                    # Save to inventory
                    df_inv = pd.concat([df_inv, new_item], ignore_index=True)
                    save_inventory(df_inv)
                    
                    # Record transaction
                    add_transaction(
                        "Receiving", 
                        item_code, 
                        item_name, 
                        total_quantity_g, 
                        unit, 
                        f"New culture media added - Lot: {lot_no}"
                    )
                    
                    st.success(f"✅ Culture Media '{item_name}' added successfully!")
                    st.info(f"📝 Item Code: {item_code}")
                    st.info(f"📊 Total Quantity: {total_quantity_g} {unit}")
                    st.info(f"🧪 For 1000mL: {per_1000ml_g} g (for calculations)")
                    st.balloons()
    
    # =========================================
    # STRAINS (Reference Materials)
    # =========================================
    elif category == "Strains (Reference Materials)":
        st.subheader("🧪 Reference Strains Entry")
        
        with st.form("strains_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                strain_name = st.text_input("Strain Name *", placeholder="e.g., Escherichia coli ATCC 25922")
                lot_no = st.text_input("Lot/Batch No. *", placeholder="e.g., 12345-AB")
                manufacturer = st.text_input("Manufacturer *", placeholder="e.g., ATCC, Microbiologics")
                
            with col2:
                storage_conditions = st.selectbox("Storage Conditions *", STORAGE_CONDITIONS)
                validity_date = st.date_input("Validity/Expiry Date *")
                quantity = st.number_input("Quantity *", min_value=1, step=1, format="%d")
                unit = st.selectbox("Unit", ["vial", "ampoule", "culture", "lyophilized"], index=0)
            
            col3, col4 = st.columns(2)
            with col3:
                passage_number = st.text_input("Passage Number", placeholder="e.g., P2")
                date_received = st.date_input("Date Received", value=date.today())
            with col4:
                notes = st.text_area("Notes", placeholder="Special characteristics, resistance patterns, etc.")
            
            st.markdown("---")
            submitted = st.form_submit_button("💾 Save Reference Strain", use_container_width=True)
            
            if submitted:
                if not strain_name or not lot_no or not manufacturer or not validity_date or quantity <= 0:
                    st.error("❌ Please fill all required fields (*)")
                else:
                    # Generate item code
                    item_code = generate_item_code(df_inv, category)
                    
                    # Prepare data for new item
                    new_item_data = [
                        item_code, strain_name, category, strain_name, lot_no,
                        manufacturer, storage_conditions, validity_date,
                        quantity, unit, passage_number, date_received, notes
                    ]
                    
                    # Create new row with all columns
                    new_item = create_new_item_row(new_item_data, STRAINS_COLUMNS)
                    
                    # Save to inventory
                    df_inv = pd.concat([df_inv, new_item], ignore_index=True)
                    save_inventory(df_inv)
                    
                    # Record transaction
                    add_transaction("Receiving", item_code, strain_name, quantity, unit, f"New strain added - Lot: {lot_no}")
                    
                    st.success(f"✅ Reference Strain '{strain_name}' added successfully!")
                    st.info(f"📝 Item Code: {item_code}")
                    st.balloons()
    
    # =========================================
    # DISPOSAL (Consumables)
    # =========================================
    elif category == "Disposal (Petri dishes, tissues, membrane filters, etc)":
        st.subheader("🧴 Disposable Items Entry")
        
        with st.form("disposal_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                item_name = st.text_input("Item Name *", placeholder="e.g., Petri dishes 90mm")
                description = st.text_area("Description", placeholder="Detailed description of the item")
                manufacturer = st.text_input("Manufacturer/Brand", placeholder="e.g., Sarstedt, Falcon")
                
            with col2:
                catalog_no = st.text_input("Catalog No.", placeholder="e.g., 82.1472")
                quantity = st.number_input("Quantity *", min_value=0, step=1, format="%d")
                unit = st.selectbox("Unit *", DISPOSAL_UNITS)
                if unit == "Other":
                    unit = st.text_input("Specify unit")
            
            col3, col4 = st.columns(2)
            with col3:
                package_size = st.text_input("Package Size", placeholder="e.g., 20 plates/sleeve")
                storage_location = st.text_input("Storage Location", placeholder="Shelf/Cabinet/Room")
            with col4:
                min_stock = st.number_input("Minimum Stock Level", min_value=0, step=1, format="%d")
                date_received = st.date_input("Date Received", value=date.today())
                notes = st.text_area("Notes", placeholder="Additional information")
            
            st.markdown("---")
            submitted = st.form_submit_button("💾 Save Disposable Item", use_container_width=True)
            
            if submitted:
                if not item_name or quantity <= 0:
                    st.error("❌ Please fill all required fields (*)")
                else:
                    # Generate item code
                    item_code = generate_item_code(df_inv, category)
                    
                    # Prepare data for new item
                    new_item_data = [
                        item_code, item_name, category, description, quantity,
                        unit, package_size, manufacturer, catalog_no, storage_location,
                        min_stock, date_received, notes
                    ]
                    
                    # Create new row with all columns
                    new_item = create_new_item_row(new_item_data, DISPOSAL_COLUMNS)
                    
                    # Save to inventory
                    df_inv = pd.concat([df_inv, new_item], ignore_index=True)
                    save_inventory(df_inv)
                    
                    # Record transaction
                    add_transaction("Receiving", item_code, item_name, quantity, unit, "New disposable item added")
                    
                    st.success(f"✅ Item '{item_name}' added successfully!")
                    st.info(f"📝 Item Code: {item_code}")
                    st.balloons()
    
    # =========================================
    # OTHER CATEGORIES
    # =========================================
    else:
        st.subheader("📦 Generic Item Entry")
        
        with st.form("generic_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                item_name = st.text_input("Item Name *")
                manufacturer = st.text_input("Manufacturer/Supplier")
                quantity = st.number_input("Quantity *", min_value=0.0, step=0.1)
                
            with col2:
                unit = st.selectbox("Unit", UNITS)
                expiry_date = st.date_input("Expiry Date", value=None)
                storage_location = st.text_input("Storage Location")
            
            notes = st.text_area("Notes")
            
            st.markdown("---")
            submitted = st.form_submit_button("💾 Save Item", use_container_width=True)
            
            if submitted:
                if not item_name or quantity <= 0:
                    st.error("❌ Please fill all required fields (*)")
                else:
                    # Generate item code
                    item_code = generate_item_code(df_inv, category)
                    
                    # Create a generic entry with available columns
                    generic_columns = ["Item Code", "Item Name", "Category", "Quantity", "Unit", 
                                      "Manufacturer", "Expiry Date", "Storage Location", "Notes"]
                    new_item_data = [
                        item_code, item_name, category, quantity, unit,
                        manufacturer, expiry_date, storage_location, notes
                    ]
                    
                    new_item = create_new_item_row(new_item_data, generic_columns)
                    
                    df_inv = pd.concat([df_inv, new_item], ignore_index=True)
                    save_inventory(df_inv)
                    add_transaction("Receiving", item_code, item_name, quantity, unit, "New item added")
                    
                    st.success(f"✅ Item '{item_name}' added successfully!")
                    st.info(f"📝 Item Code: {item_code}")
                    st.balloons()

# =====================================================
# Prepare Media Section
# =====================================================
elif menu == "Prepare Media":
    st.header("🧪 Prepare Culture Media")
    
    df_inv = load_inventory()
    
    # Filter only culture media items
    media_items = df_inv[df_inv["Category"] == "Culture Media (Broths, Agars, etc)"]
    
    if media_items.empty:
        st.warning("⚠️ No culture media found in inventory. Please add culture media first.")
        st.stop()
    
    st.markdown("""
    This section helps you calculate how much media powder you need based on your desired volume.
    The calculation uses the 'For 1000mL' value you entered when adding the media.
    """)
    
    # Select media to use
    col1, col2 = st.columns(2)
    
    with col1:
        selected_media = st.selectbox(
            "Select Media to Prepare",
            media_items["Item Name"].tolist()
        )
    
    # Get selected media details
    media_data = media_items[media_items["Item Name"] == selected_media].iloc[0]
    
    with col2:
        st.info(f"""
        **Media Details:**
        - Item Code: {media_data['Item Code']}
        - Lot No.: {media_data['Lot No.']}
        - Available in Inventory: {media_data['Current Quantity (g)']} g
        - For 1000mL (calculation reference): {media_data['For 1000mL (g)']} g
        """)
    
    st.markdown("---")
    
    with st.form("media_prep_form"):
        col3, col4 = st.columns(2)
        
        with col3:
            st.subheader("📐 Calculate Media Needed")
            volume_to_prepare = st.number_input(
                "Volume to Prepare (mL) *",
                min_value=10,
                max_value=10000,
                value=1000,
                step=100,
                help="Enter the volume of media you want to prepare"
            )
            
            # Calculate media needed based on the For 1000mL reference value
            per_1000ml = float(media_data['For 1000mL (g)'])
            media_needed = calculate_media_needed(volume_to_prepare, per_1000ml)
            
            st.metric(
                "Media Powder Needed", 
                f"{media_needed:.2f} g",
                help=f"Based on {per_1000ml}g per 1000mL"
            )
            
            # Show calculation formula
            st.caption(f"Calculation: ({volume_to_prepare} mL ÷ 1000) × {per_1000ml} g = {media_needed:.2f} g")
        
        with col4:
            st.subheader("💧 Water Required")
            water_volume = st.number_input(
                "Distilled Water Volume (mL)",
                value=volume_to_prepare,
                disabled=True,
                help="Typically equal to the volume you're preparing"
            )
            
            # Check if enough media available
            current_qty = float(media_data['Current Quantity (g)'])
            if media_needed > current_qty:
                st.error(f"❌ INSUFFICIENT MEDIA!")
                st.error(f"Available: {current_qty:.2f} g")
                st.error(f"Needed: {media_needed:.2f} g")
                st.error(f"Shortage: {media_needed - current_qty:.2f} g")
            else:
                remaining = current_qty - media_needed
                st.success(f"✅ Sufficient media available")
                st.info(f"📊 Will remain after preparation: {remaining:.2f} g")
        
        st.markdown("---")
        st.subheader("📝 Preparation Details")
        
        col5, col6 = st.columns(2)
        with col5:
            prep_date = st.date_input("Preparation Date", value=date.today())
            prepared_by = st.text_input("Prepared By *")
        
        with col6:
            media_expiry = st.date_input(
                "Prepared Media Expiry Date",
                value=date.today() + timedelta(days=7),
                help="Typically 1-4 weeks depending on media type and storage"
            )
            sterilization_method = st.selectbox(
                "Sterilization Method",
                ["Autoclave (121°C for 15 min)", "Filtration (0.22µm)", "Boiling", "None Required", "Other"]
            )
        
        notes = st.text_area("Notes", placeholder="pH adjustment, special additives, etc.")
        
        submitted = st.form_submit_button("✅ Record Media Preparation", use_container_width=True)
        
        if submitted:
            if not prepared_by:
                st.error("❌ Please enter who prepared the media")
            elif media_needed > current_qty:
                st.error("❌ Cannot prepare: Insufficient media available")
            else:
                # Update inventory (subtract the amount used)
                new_qty = current_qty - media_needed
                df_inv.loc[df_inv["Item Code"] == media_data["Item Code"], "Current Quantity (g)"] = new_qty
                save_inventory(df_inv)
                
                # Record transaction
                add_transaction(
                    "Issuing",
                    media_data["Item Code"],
                    selected_media,
                    media_needed,
                    "g",
                    f"Media preparation - {volume_to_prepare}mL"
                )
                
                # Record media preparation
                df_prep = load_media_prep()
                new_prep = pd.DataFrame([[
                    prep_date,
                    selected_media,
                    media_data["Lot No."],
                    volume_to_prepare,
                    media_needed,
                    water_volume,
                    prepared_by,
                    media_expiry,
                    sterilization_method,
                    notes
                ]], columns=["Date", "Media Type", "Lot No.", "Quantity (mL)", 
                            "Media Used (g)", "Water Used (mL)", "Prepared By", 
                            "Expiry Date", "Sterilization Method", "Notes"])
                
                df_prep = pd.concat([df_prep, new_prep], ignore_index=True)
                save_media_prep(df_prep)
                
                st.success(f"✅ Successfully prepared {volume_to_prepare}mL of {selected_media}")
                st.info(f"📊 Media used: {media_needed:.2f} g")
                st.info(f"📦 Remaining in inventory: {new_qty:.2f} g")
                
                # Show preparation summary
                with st.expander("📋 Preparation Summary"):
                    st.write(f"**Media:** {selected_media}")
                    st.write(f"**Lot Number:** {media_data['Lot No.']}")
                    st.write(f"**Volume Prepared:** {volume_to_prepare} mL")
                    st.write(f"**Media Used:** {media_needed:.2f} g")
                    st.write(f"**Water Used:** {water_volume} mL")
                    st.write(f"**Sterilization:** {sterilization_method}")
                    st.write(f"**Prepared by:** {prepared_by}")
                    st.write(f"**Date:** {prep_date}")
                    st.write(f"**Expiry:** {media_expiry}")
                    if notes:
                        st.write(f"**Notes:** {notes}")
                
                st.balloons()

# =====================================================
# Issue Material
# =====================================================
elif menu == "Issue Material":
    st.header("📤 Issue Material from Store")
    
    df_inv = load_inventory()
    
    if df_inv.empty:
        st.warning("⚠️ No items in inventory")
        st.stop()
    
    # Category filter
    categories = ["All"] + df_inv["Category"].unique().tolist()
    selected_category = st.selectbox("Filter by Category", categories)
    
    if selected_category != "All":
        filtered_items = df_inv[df_inv["Category"] == selected_category]
    else:
        filtered_items = df_inv
    
    # Select item
    item_names = filtered_items["Item Name"].tolist()
    if not item_names:
        st.warning("No items in selected category")
        st.stop()
    
    selected_item = st.selectbox("Select Item to Issue", item_names)
    
    # Get item details
    item_data = filtered_items[filtered_items["Item Name"] == selected_item].iloc[0]
    
    # Display current quantity based on category
    col_info1, col_info2 = st.columns(2)
    
    with col_info1:
        if item_data["Category"] == "Culture Media (Broths, Agars, etc)":
            current_qty = float(item_data.get("Current Quantity (g)", 0))
            unit = "g"
            st.info(f"📦 Available: **{current_qty} {unit}**")
        elif item_data["Category"] == "Strains (Reference Materials)":
            current_qty = float(item_data.get("Quantity", 0))
            unit = item_data.get("Unit", "vial")
            st.info(f"📦 Available: **{current_qty} {unit}(s)**")
            if pd.notna(item_data.get("Passage Number", "")):
                st.info(f"🧬 Passage: **{item_data['Passage Number']}**")
        else:
            current_qty = float(item_data.get("Quantity", 0))
            unit = item_data.get("Unit", "unit")
            st.info(f"📦 Available: **{current_qty} {unit}(s)**")
    
    with col_info2:
        if pd.notna(item_data.get("Storage Conditions", "")):
            st.info(f"❄️ Storage: **{item_data['Storage Conditions']}**")
        if pd.notna(item_data.get("Storage Location", "")):
            st.info(f"📍 Location: **{item_data['Storage Location']}**")
    
    st.markdown("---")
    
    with st.form("issue_form"):
        col_issue1, col_issue2 = st.columns(2)
        
        with col_issue1:
            issue_qty = st.number_input(
                f"Quantity to Issue ({unit})",
                min_value=0.0,
                max_value=float(current_qty),
                step=0.1 if unit in ["g", "mL"] else 1,
                format="%.2f" if unit in ["g", "mL"] else "%d"
            )
        
        with col_issue2:
            purpose = st.selectbox(
                "Purpose",
                ["Testing", "Media Preparation", "Quality Control", "Research", "Teaching", "Maintenance", "Other"]
            )
            if purpose == "Other":
                purpose = st.text_input("Specify purpose")
        
        recipient = st.text_input("Recipient Name *")
        notes = st.text_area("Additional Notes")
        
        submitted = st.form_submit_button("✅ Confirm Issue", use_container_width=True)
        
        if submitted:
            if issue_qty <= 0:
                st.error("❌ Quantity must be greater than zero")
            elif issue_qty > current_qty:
                st.error(f"❌ Requested quantity exceeds available")
            elif not recipient:
                st.warning("⚠️ Please enter recipient name")
            else:
                # Update quantity based on category
                if item_data["Category"] == "Culture Media (Broths, Agars, etc)":
                    new_qty = current_qty - issue_qty
                    df_inv.loc[df_inv["Item Code"] == item_data["Item Code"], "Current Quantity (g)"] = new_qty
                else:
                    new_qty = current_qty - issue_qty
                    df_inv.loc[df_inv["Item Code"] == item_data["Item Code"], "Quantity"] = new_qty
                
                df_inv.loc[df_inv["Item Code"] == item_data["Item Code"], "Last Updated"] = date.today()
                save_inventory(df_inv)
                
                # Record transaction
                reason_text = f"Issued for {purpose}"
                if notes:
                    reason_text += f" - {notes}"
                add_transaction(
                    "Issuing",
                    item_data["Item Code"],
                    selected_item,
                    issue_qty,
                    unit,
                    reason_text
                )
                
                st.success(f"✅ Successfully issued {issue_qty} {unit} of {selected_item} to {recipient}")
                
                # Check if low stock
                if pd.notna(item_data.get("Min Stock Level", "")):
                    min_stock = float(item_data["Min Stock Level"])
                    if new_qty <= min_stock:
                        st.warning(f"⚠️ Warning: Remaining quantity ({new_qty} {unit}) is at or below minimum stock level ({min_stock} {unit})")

# =====================================================
# Dashboard
# =====================================================
elif menu == "Dashboard":
    st.header("📊 Microbiology Lab Dashboard")
    
    df_inv = load_inventory()
    df_trans = load_transactions()
    df_prep = load_media_prep()
    
    if df_inv.empty:
        st.info("📝 No items in inventory yet. Start by adding materials.")
    else:
        # Key metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Items", len(df_inv))
        
        with col2:
            media_count = len(df_inv[df_inv["Category"] == "Culture Media (Broths, Agars, etc)"])
            st.metric("Culture Media", media_count)
        
        with col3:
            strains_count = len(df_inv[df_inv["Category"] == "Strains (Reference Materials)"])
            st.metric("Reference Strains", strains_count)
        
        with col4:
            disposables = len(df_inv[df_inv["Category"].str.contains("Disposal", na=False)])
            st.metric("Disposables", disposables)
        
        st.markdown("---")
        
        # Simple pie chart with matplotlib
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.subheader("Items by Category")
            category_dist = df_inv["Category"].value_counts()
            
            if not category_dist.empty:
                # Create pie chart
                fig, ax = plt.subplots(figsize=(8, 6))
                colors = plt.cm.Set3(range(len(category_dist)))
                
                wedges, texts, autotexts = ax.pie(
                    category_dist.values,
                    labels=category_dist.index,
                    autopct='%1.1f%%',
                    colors=colors,
                    startangle=90
                )
                ax.set_title("Distribution by Category")
                st.pyplot(fig)
            else:
                st.info("No category data available")
        
        with col_chart2:
            st.subheader("Recent Media Preparations")
            if not df_prep.empty:
                last_5 = df_prep.sort_values("Date", ascending=False).head(5)
                st.dataframe(
                    last_5[["Date", "Media Type", "Quantity (mL)", "Prepared By"]],
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No media preparations recorded")
        
        st.markdown("---")
        
        # Expiring soon items
        st.subheader("⚠️ Items Expiring Soon")
        
        expiring_items = []
        for _, row in df_inv.iterrows():
            expiry_col = None
            if "Expiry Date" in row.index and pd.notna(row["Expiry Date"]):
                expiry_col = "Expiry Date"
            elif "Validity Date" in row.index and pd.notna(row["Validity Date"]):
                expiry_col = "Validity Date"
            
            if expiry_col:
                try:
                    expiry_date = pd.to_datetime(row[expiry_col])
                    days_left = (expiry_date - pd.Timestamp.now()).days
                    if 0 <= days_left <= 30:
                        expiring_items.append({
                            "Item": row["Item Name"],
                            "Category": row["Category"],
                            "Expiry Date": expiry_date.strftime("%Y-%m-%d"),
                            "Days Left": days_left
                        })
                except:
                    pass
        
        if expiring_items:
            expiring_df = pd.DataFrame(expiring_items)
            st.dataframe(expiring_df, use_container_width=True, hide_index=True)
        else:
            st.success("No items expiring in the next 30 days")

# =====================================================
# Inventory (View all items)
# =====================================================
elif menu == "Inventory":
    st.header("📋 Complete Inventory")
    
    df_inv = load_inventory()
    
    if df_inv.empty:
        st.info("No items in inventory")
        st.stop()
    
    # Category filter
    categories = ["All"] + df_inv["Category"].unique().tolist()
    selected_category = st.selectbox("Filter by Category", categories)
    
    if selected_category != "All":
        filtered_df = df_inv[df_inv["Category"] == selected_category]
    else:
        filtered_df = df_inv
    
    # Display
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    
    # Export button
    if st.button("📥 Export to Excel"):
        output_file = f"inventory_export_{date.today().strftime('%Y%m%d')}.xlsx"
        filtered_df.to_excel(output_file, index=False)
        with open(output_file, "rb") as f:
            st.download_button(
                "Download Excel File",
                data=f.read(),
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =====================================================
# Transactions
# =====================================================
elif menu == "Transactions":
    st.header("📊 Transaction History")
    
    df_trans = load_transactions()
    
    if df_trans.empty:
        st.info("No transactions recorded")
        st.stop()
    
    # Date filter
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("From", value=date.today() - timedelta(days=30))
    with col2:
        end_date = st.date_input("To", value=date.today())
    
    # Filter
    df_trans["Date"] = pd.to_datetime(df_trans["Date"])
    mask = (df_trans["Date"] >= pd.Timestamp(start_date)) & (df_trans["Date"] <= pd.Timestamp(end_date))
    filtered = df_trans[mask]
    
    # Summary
    st.metric("Transactions in period", len(filtered))
    
    # Display
    filtered["Date"] = filtered["Date"].dt.strftime("%Y-%m-%d")
    st.dataframe(filtered, use_container_width=True, hide_index=True)

# =====================================================
# Reports
# =====================================================
elif menu == "Reports":
    st.header("📈 Reports")
    
    report_type = st.selectbox(
        "Select Report",
        ["Media Usage Report", "Strains Inventory", "Low Stock Alert", "Preparation Log"]
    )
    
    if report_type == "Media Usage Report":
        st.subheader("Media Usage Report")
        df_prep = load_media_prep()
        if not df_prep.empty:
            # Summary by media type
            summary = df_prep.groupby("Media Type").agg({
                "Quantity (mL)": "sum",
                "Media Used (g)": "sum"
            }).reset_index()
            
            col1, col2 = st.columns(2)
            with col1:
                # Bar chart with matplotlib
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.bar(summary["Media Type"], summary["Quantity (mL)"])
                ax.set_xlabel("Media Type")
                ax.set_ylabel("Total Volume (mL)")
                ax.set_title("Total Volume Prepared by Media Type")
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                st.pyplot(fig)
            
            with col2:
                # Bar chart for media used
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.bar(summary["Media Type"], summary["Media Used (g)"])
                ax.set_xlabel("Media Type")
                ax.set_ylabel("Media Used (g)")
                ax.set_title("Total Media Used (g)")
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                st.pyplot(fig)
            
            st.dataframe(df_prep, use_container_width=True, hide_index=True)
        else:
            st.info("No media preparation records")
    
    elif report_type == "Strains Inventory":
        st.subheader("Reference Strains Inventory")
        df_inv = load_inventory()
        strains = df_inv[df_inv["Category"] == "Strains (Reference Materials)"]
        
        if not strains.empty:
            st.dataframe(strains, use_container_width=True, hide_index=True)
        else:
            st.info("No strains in inventory")
    
    elif report_type == "Low Stock Alert":
        st.subheader("Low Stock Items")
        df_inv = load_inventory()
        
        low_stock_items = []
        for _, row in df_inv.iterrows():
            if "Min Stock Level" in row.index and pd.notna(row["Min Stock Level"]):
                current = float(row.get("Quantity", 0)) if "Quantity" in row.index else float(row.get("Current Quantity (g)", 0))
                min_stock = float(row["Min Stock Level"])
                if current <= min_stock:
                    low_stock_items.append({
                        "Item": row["Item Name"],
                        "Category": row["Category"],
                        "Current": current,
                        "Unit": row.get("Unit", "unit"),
                        "Min Stock": min_stock
                    })
        
        if low_stock_items:
            st.dataframe(pd.DataFrame(low_stock_items), use_container_width=True, hide_index=True)
        else:
            st.success("No low stock items found")
    
    elif report_type == "Preparation Log":
        st.subheader("Media Preparation Log")
        df_prep = load_media_prep()
        if not df_prep.empty:
            st.dataframe(df_prep, use_container_width=True, hide_index=True)
        else:
            st.info("No preparation records")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray; padding: 10px;'>
    Microbiology Lab Store Management System v1.0 © 2026
    </div>
    """,
    unsafe_allow_html=True
)