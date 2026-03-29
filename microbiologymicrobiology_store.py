import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
import os
import re
import io

# =====================================================
# Configuration
# =====================================================
MEDIA_MASTER_FILE = "media_master.xlsx"  # Master list of all media (added once)
MEDIA_BATCH_FILE = "media_batches.xlsx"  # Batch records (each preparation)
DISPOSABLE_MASTER_FILE = "disposable_master.xlsx"
STRAIN_MASTER_FILE = "strain_master.xlsx"

# =====================================================
# Page Configuration
# =====================================================
st.set_page_config(page_title="Microbiology Store", page_icon="🏪", layout="wide")

st.title("🏪 Microbiology Store - Inventory Management")
st.markdown("---")

# =====================================================
# Load/Save Functions
# =====================================================

def load_media_master():
    """Load master list of all media (added once)"""
    if os.path.exists(MEDIA_MASTER_FILE):
        df = pd.read_excel(MEDIA_MASTER_FILE)
        return df
    return pd.DataFrame(columns=[
        "Media ID", "Media Type", "Lot Number", "Reference Number", 
        "Whole Quantity", "Unit", "Expiry Date", "Open Date",
        "Grams_per_ml", "Distilled_Water_ml", "Batch_Prefix"
    ])

def save_media_master(df):
    """Save master list of all media"""
    df.to_excel(MEDIA_MASTER_FILE, index=False)
    st.success("✅ Media master data saved!")

def load_media_batches():
    """Load batch preparation records"""
    if os.path.exists(MEDIA_BATCH_FILE):
        df = pd.read_excel(MEDIA_BATCH_FILE)
        if "Consumed_Quantity" not in df.columns:
            df["Consumed_Quantity"] = 0.0
        if "Remaining_Quantity" not in df.columns:
            df["Remaining_Quantity"] = df["Prepared_Quantity"]
        return df
    return pd.DataFrame(columns=[
        "Batch_ID", "Media_ID", "Media_Type", "Batch_Number", 
        "Preparation_Date", "Prepared_Quantity", "Unit", 
        "Expiry_Date", "Prepared_By", "Consumed_Quantity", 
        "Remaining_Quantity", "Status"
    ])

def save_media_batches(df):
    """Save batch preparation records"""
    df.to_excel(MEDIA_BATCH_FILE, index=False)
    st.success("✅ Batch data saved!")

def generate_batch_number(prefix, media_id, year):
    """
    Generate batch number: PREFIX-SERIAL-YEAR
    Example: TSA-001-2026
    """
    existing = load_media_batches()
    if not existing.empty:
        # Filter by prefix and year
        pattern = f"{prefix}-(\\d+)-{year}"
        existing_serials = []
        for batch in existing["Batch_Number"]:
            match = re.search(pattern, str(batch))
            if match:
                existing_serials.append(int(match.group(1)))
        
        if existing_serials:
            next_serial = max(existing_serials) + 1
        else:
            next_serial = 1
    else:
        next_serial = 1
    
    return f"{prefix}-{next_serial:03d}-{year}"

# =====================================================
# Navigation Tabs
# =====================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📦 Add Media (Master)", 
    "🧪 Prepare Media (Batch)", 
    "📋 Media Batches Inventory",
    "📊 Consumption Log",
    "⚠️ Expiry Alerts"
])

# =====================================================
# TAB 1: Add Media to Master (Added Once)
# =====================================================
with tab1:
    st.subheader("📦 Add New Media to Master List")
    st.caption("Add each media type once. This defines the media properties.")
    
    with st.form("add_media_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            media_type = st.text_input("Media Type *", placeholder="e.g., TSA, SDA, FTM, TSB")
            lot_number = st.text_input("Lot Number (Manufacturer)", placeholder="Manufacturer's lot number")
            reference_number = st.text_input("Reference Number/Catalog #", placeholder="e.g., CM0123")
            
            whole_quantity = st.number_input("Whole Quantity *", min_value=0.0, value=500.0, step=50.0)
            unit = st.selectbox("Unit *", ["mL", "g", "L", "kg", "plates", "bottles"])
            
        with col2:
            expiry_date = st.date_input("Expiry Date (Master)", value=date.today() + timedelta(days=365))
            open_date = st.date_input("Open Date", value=date.today())
            
            st.markdown("### 📐 Preparation Formula")
            grams = st.number_input("How many grams?", min_value=0.0, value=40.0, step=5.0)
            water_ml = st.number_input("For how many mL distilled water?", min_value=0.0, value=1000.0, step=100.0)
            
            st.markdown("### 🏷️ Batch Number Rule")
            batch_prefix = st.text_input("Batch Prefix *", placeholder="e.g., TSA, SDA, FTM", help="Will be used to generate batch numbers: PREFIX-SERIAL-YEAR")
        
        submitted = st.form_submit_button("✅ Add Media to Master", type="primary")
        
        if submitted:
            if not media_type:
                st.error("❌ Media Type is required!")
            elif not batch_prefix:
                st.error("❌ Batch Prefix is required!")
            elif whole_quantity <= 0:
                st.error("❌ Whole Quantity must be greater than 0!")
            else:
                # Generate Media ID
                existing = load_media_master()
                media_id = f"MED-{len(existing) + 1:04d}"
                
                new_media = pd.DataFrame([{
                    "Media ID": media_id,
                    "Media Type": media_type,
                    "Lot Number": lot_number,
                    "Reference Number": reference_number,
                    "Whole Quantity": whole_quantity,
                    "Unit": unit,
                    "Expiry Date": expiry_date,
                    "Open Date": open_date,
                    "Grams_per_ml": grams,
                    "Distilled_Water_ml": water_ml,
                    "Batch_Prefix": batch_prefix.upper()
                }])
                
                updated = pd.concat([existing, new_media], ignore_index=True)
                save_media_master(updated)
                st.success(f"✅ Media added successfully! Media ID: **{media_id}**")
                st.balloons()

    # Display existing media master
    st.markdown("---")
    st.subheader("📋 Existing Media Master List")
    media_master = load_media_master()
    if not media_master.empty:
        st.dataframe(media_master, use_container_width=True)
        
        # Option to delete media
        with st.expander("🗑️ Delete Media from Master"):
            media_to_delete = st.selectbox("Select Media to Delete", media_master["Media Type"].tolist())
            if st.button("Delete Selected Media", type="secondary"):
                media_master = media_master[media_master["Media Type"] != media_to_delete]
                save_media_master(media_master)
                st.rerun()
    else:
        st.info("No media added yet. Use the form above to add media.")

# =====================================================
# TAB 2: Prepare Media (Create Batch)
# =====================================================
with tab2:
    st.subheader("🧪 Prepare New Media Batch")
    st.caption("Each preparation creates a new batch with unique batch number")
    
    media_master = load_media_master()
    
    if media_master.empty:
        st.warning("⚠️ No media found in master list. Please add media in 'Add Media' tab first.")
    else:
        with st.form("prepare_batch_form"):
            # Select media from master
            media_options = media_master["Media Type"].tolist()
            selected_media = st.selectbox("Select Media to Prepare", media_options)
            
            # Get selected media details
            media_row = media_master[media_master["Media Type"] == selected_media].iloc[0]
            batch_prefix = media_row["Batch_Prefix"]
            unit = media_row["Unit"]
            grams_per_ml = media_row["Grams_per_ml"]
            water_ml = media_row["Distilled_Water_ml"]
            
            st.info(f"""
            **Media Information:**
            - Batch Prefix: `{batch_prefix}`
            - Preparation Formula: `{grams_per_ml} g` for `{water_ml} mL` distilled water
            - Unit: {unit}
            """)
            
            col1, col2 = st.columns(2)
            
            with col1:
                prepared_quantity = st.number_input(f"Prepared Quantity ({unit})", min_value=0.0, value=500.0, step=50.0)
                prepared_by = st.text_input("Prepared By", value=st.session_state.get("username", "Lab Technician"))
            
            with col2:
                preparation_date = st.date_input("Preparation Date", value=date.today())
                # Calculate expiry (e.g., 30 days from preparation)
                expiry_date = st.date_input("Expiry Date", value=preparation_date + timedelta(days=30))
            
            submitted = st.form_submit_button("✅ Create Batch", type="primary")
            
            if submitted:
                if prepared_quantity <= 0:
                    st.error("❌ Prepared Quantity must be greater than 0!")
                else:
                    # Generate batch number
                    year = preparation_date.year
                    batch_number = generate_batch_number(batch_prefix, media_row["Media ID"], year)
                    
                    # Create batch record
                    batches = load_media_batches()
                    batch_id = f"BATCH-{len(batches) + 1:06d}"
                    
                    new_batch = pd.DataFrame([{
                        "Batch_ID": batch_id,
                        "Media_ID": media_row["Media ID"],
                        "Media_Type": selected_media,
                        "Batch_Number": batch_number,
                        "Preparation_Date": preparation_date,
                        "Prepared_Quantity": prepared_quantity,
                        "Unit": unit,
                        "Expiry_Date": expiry_date,
                        "Prepared_By": prepared_by,
                        "Consumed_Quantity": 0.0,
                        "Remaining_Quantity": prepared_quantity,
                        "Status": "Active"
                    }])
                    
                    updated = pd.concat([batches, new_batch], ignore_index=True)
                    save_media_batches(updated)
                    st.success(f"""
                    ✅ Batch created successfully!
                    
                    **Batch Number:** `{batch_number}`
                    **Media:** {selected_media}
                    **Quantity:** {prepared_quantity} {unit}
                    **Expiry:** {expiry_date}
                    """)
                    st.balloons()

# =====================================================
# TAB 3: Media Batches Inventory
# =====================================================
with tab3:
    st.subheader("📋 All Media Batches")
    
    batches = load_media_batches()
    
    if not batches.empty:
        # Filters
        col1, col2, col3 = st.columns(3)
        with col1:
            media_filter = st.multiselect("Filter by Media Type", options=batches["Media_Type"].unique().tolist())
        with col2:
            status_filter = st.multiselect("Filter by Status", options=batches["Status"].unique().tolist())
        with col3:
            show_expired = st.checkbox("Show Expired Batches", value=False)
        
        filtered = batches.copy()
        if media_filter:
            filtered = filtered[filtered["Media_Type"].isin(media_filter)]
        if status_filter:
            filtered = filtered[filtered["Status"].isin(status_filter)]
        if not show_expired:
            today = pd.Timestamp(date.today())
            filtered = filtered[pd.to_datetime(filtered["Expiry_Date"]) >= today]
        
        # Display
        display_df = filtered.copy()
        display_df["Preparation_Date"] = pd.to_datetime(display_df["Preparation_Date"]).dt.strftime("%Y-%m-%d")
        display_df["Expiry_Date"] = pd.to_datetime(display_df["Expiry_Date"]).dt.strftime("%Y-%m-%d")
        
        st.dataframe(display_df, use_container_width=True)
        
        # Summary
        st.markdown("---")
        st.subheader("📊 Batch Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_batches = len(filtered)
            st.metric("Total Active Batches", total_batches)
        with col2:
            total_remaining = filtered["Remaining_Quantity"].sum() if not filtered.empty else 0
            st.metric("Total Remaining", f"{total_remaining:.0f} {filtered['Unit'].iloc[0] if not filtered.empty else ''}")
        with col3:
            total_consumed = filtered["Consumed_Quantity"].sum() if not filtered.empty else 0
            st.metric("Total Consumed", f"{total_consumed:.0f} {filtered['Unit'].iloc[0] if not filtered.empty else ''}")
        with col4:
            utilization = (total_consumed / (total_consumed + total_remaining) * 100) if (total_consumed + total_remaining) > 0 else 0
            st.metric("Overall Utilization", f"{utilization:.1f}%")
        
        # Download
        output = io.BytesIO()
        export_df = filtered.copy()
        export_df["Preparation_Date"] = pd.to_datetime(export_df["Preparation_Date"]).dt.strftime("%Y-%m-%d")
        export_df["Expiry_Date"] = pd.to_datetime(export_df["Expiry_Date"]).dt.strftime("%Y-%m-%d")
        export_df.to_excel(output, index=False, engine="openpyxl")
        st.download_button("📥 Download Batches as Excel", data=output.getvalue(), 
                          file_name="media_batches.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No batches prepared yet. Go to 'Prepare Media' tab to create batches.")

# =====================================================
# TAB 4: Consumption Log
# =====================================================
with tab4:
    st.subheader("📊 Consumption Log")
    
    batches = load_media_batches()
    
    if not batches.empty:
        consumed = batches[batches["Consumed_Quantity"] > 0].copy()
        
        if not consumed.empty:
            consumed["Consumption_Date"] = pd.to_datetime(consumed.get("Last_Consumption_Date", pd.NaT), errors="coerce")
            consumed["Utilization_%"] = (consumed["Consumed_Quantity"] / consumed["Prepared_Quantity"]) * 100
            
            st.dataframe(consumed[["Batch_Number", "Media_Type", "Prepared_Quantity", "Consumed_Quantity", 
                                   "Utilization_%", "Status"]], use_container_width=True)
            
            # Chart
            st.markdown("---")
            st.subheader("Consumption by Media Type")
            consumption_by_type = consumed.groupby("Media_Type")["Consumed_Quantity"].sum().sort_values(ascending=False)
            
            if not consumption_by_type.empty:
                import matplotlib.pyplot as plt
                fig, ax = plt.subplots(figsize=(10, 5))
                consumption_by_type.plot(kind="bar", ax=ax, color="skyblue")
                ax.set_xlabel("Media Type")
                ax.set_ylabel("Quantity Consumed")
                ax.set_title("Total Media Consumption by Type")
                ax.tick_params(axis="x", rotation=45)
                st.pyplot(fig)
        else:
            st.info("No consumption records yet. Consumption happens when you perform tests in the main app.")
    else:
        st.info("No data available.")

# =====================================================
# TAB 5: Expiry Alerts
# =====================================================
with tab5:
    st.subheader("⚠️ Expiry Alerts")
    
    batches = load_media_batches()
    
    if not batches.empty:
        today = pd.Timestamp(date.today())
        thirty_days = today + timedelta(days=30)
        
        batches["Expiry_Date"] = pd.to_datetime(batches["Expiry_Date"])
        
        # Expired
        expired = batches[batches["Expiry_Date"] < today]
        # Expiring soon
        expiring_soon = batches[(batches["Expiry_Date"] >= today) & (batches["Expiry_Date"] <= thirty_days)]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### ❌ Expired Batches")
            if not expired.empty:
                st.dataframe(expired[["Batch_Number", "Media_Type", "Remaining_Quantity", "Unit", "Expiry_Date"]], 
                            use_container_width=True)
                
                if st.button("🗑️ Mark Expired Batches as Inactive"):
                    for idx in expired.index:
                        batches.at[idx, "Status"] = "Expired"
                    save_media_batches(batches)
                    st.rerun()
            else:
                st.success("No expired batches!")
        
        with col2:
            st.markdown("### ⏰ Expiring Within 30 Days")
            if not expiring_soon.empty:
                st.dataframe(expiring_soon[["Batch_Number", "Media_Type", "Remaining_Quantity", "Unit", "Expiry_Date"]], 
                            use_container_width=True)
            else:
                st.info("No batches expiring within 30 days.")
        
        # Waste summary
        st.markdown("---")
        total_wasted = expired["Remaining_Quantity"].sum() if not expired.empty else 0
        if total_wasted > 0:
            st.warning(f"⚠️ Total wasted volume from expired batches: **{total_wasted:.0f}** units")
    else:
        st.info("No batches available.")

# =====================================================
# Sidebar
# =====================================================
st.sidebar.title("🏪 Microbiology Store")
st.sidebar.markdown("---")
st.sidebar.info(
    """
    **Workflow:**
    
    1. **Add Media** (once per media type)
       - Define media properties
       - Set preparation formula
       - Define batch prefix
    
    2. **Prepare Media** (each time)
       - Select media from master
       - Enter prepared quantity
       - System generates batch number
    
    3. **Main App Integration**
       - Batches appear in Perform Test
       - Consumption deducted automatically
    
    **Shared Files:**
    - `media_master.xlsx`
    - `media_batches.xlsx`
    """
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.rerun()
