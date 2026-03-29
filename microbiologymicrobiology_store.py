import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
import os
import io

# =====================================================
# Configuration
# =====================================================
MEDIA_PREP_FILE = "media_preparation.xlsx"
MEDIA_INVENTORY_FILE = "microbiology_inventory.xlsx"

# =====================================================
# Page Configuration
# =====================================================
st.set_page_config(page_title="Microbiology Store", page_icon="🏪", layout="wide")

st.title("🏪 Microbiology Store - Inventory Management")
st.markdown("---")

# =====================================================
# Load/Save Functions
# =====================================================

def load_media_prep():
    """Load media preparation records"""
    if os.path.exists(MEDIA_PREP_FILE):
        df = pd.read_excel(MEDIA_PREP_FILE)
        if "Volume Consumed (mL)" not in df.columns:
            df["Volume Consumed (mL)"] = 0.0
        df["Volume Consumed (mL)"] = pd.to_numeric(df["Volume Consumed (mL)"], errors="coerce").fillna(0.0)
        df["Quantity (mL)"] = pd.to_numeric(df["Quantity (mL)"], errors="coerce").fillna(0.0)
        df["Volume Remaining (mL)"] = df["Quantity (mL)"] - df["Volume Consumed (mL)"]
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        if "Expiry Date" in df.columns:
            df["Expiry Date"] = pd.to_datetime(df["Expiry Date"], errors="coerce")
        return df
    return pd.DataFrame(columns=[
        "Date", "Media Type", "Lot No.", "Quantity (mL)", "Media Used (g)",
        "Water Used (mL)", "Prepared By", "Expiry Date", "Sterilization Method",
        "Notes", "Volume Consumed (mL)", "Volume Remaining (mL)"
    ])

def save_media_prep(df):
    """Save media preparation records"""
    save_df = df.drop(columns=["Volume Remaining (mL)"], errors="ignore")
    save_df.to_excel(MEDIA_PREP_FILE, index=False)
    st.success("✅ Media preparation data saved!")

def generate_lot_number(media_type, date_obj):
    """Generate a unique lot number: MED-YYYYMMDD-XXX"""
    date_str = date_obj.strftime("%Y%m%d")
    existing = load_media_prep()
    if not existing.empty:
        matching = existing[existing["Lot No."].str.startswith(f"MED-{date_str}-", na=False)]
        if not matching.empty:
            last_num = max([int(lot.split("-")[-1]) for lot in matching["Lot No."]])
            next_num = last_num + 1
        else:
            next_num = 1
    else:
        next_num = 1
    return f"MED-{date_str}-{next_num:03d}"

# =====================================================
# Navigation Tabs
# =====================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📝 Prepare Media", 
    "📋 Media Inventory", 
    "📊 Consumption Log", 
    "⚠️ Expiry Alerts"
])

# =====================================================
# TAB 1: Prepare Media
# =====================================================
with tab1:
    st.subheader("📝 Prepare New Media Batch")
    
    col1, col2 = st.columns(2)
    
    with col1:
        media_type = st.selectbox(
            "Media Type",
            ["TSA (Tryptone Soya Agar)",
             "SDA (Sabouraud Dextrose Agar)",
             "TSB (Tryptone Soya Broth)",
             "FTM (Fluid Thioglycollate Medium)",
             "SCDM (Soybean Casein Digest Medium)",
             "R2A Agar",
             "MacConkey Agar",
             "Blood Agar",
             "Other"]
        )
        
        if media_type == "Other":
            media_type = st.text_input("Specify Media Type")
        
        quantity_ml = st.number_input("Quantity Prepared (mL)", min_value=0.0, value=1000.0, step=100.0)
        media_used_g = st.number_input("Media Used (g)", min_value=0.0, value=40.0, step=5.0)
        water_used_ml = st.number_input("Water Used (mL)", min_value=0.0, value=1000.0, step=100.0)
    
    with col2:
        prepared_by = st.text_input("Prepared By", value=st.session_state.get("username", ""))
        expiry_date = st.date_input("Expiry Date", value=date.today() + timedelta(days=30))
        sterilization = st.selectbox("Sterilization Method", ["Autoclave", "Filtration", "Dry Heat", "Other"])
        notes = st.text_area("Notes (Optional)", placeholder="Batch specific notes...")
    
    if st.button("✅ Add Media Batch", type="primary", use_container_width=True):
        if media_type and quantity_ml > 0:
            lot_no = generate_lot_number(media_type, date.today())
            new_record = pd.DataFrame([{
                "Date": date.today(),
                "Media Type": media_type,
                "Lot No.": lot_no,
                "Quantity (mL)": quantity_ml,
                "Media Used (g)": media_used_g,
                "Water Used (mL)": water_used_ml,
                "Prepared By": prepared_by,
                "Expiry Date": expiry_date,
                "Sterilization Method": sterilization,
                "Notes": notes,
                "Volume Consumed (mL)": 0.0
            }])
            
            existing = load_media_prep()
            updated = pd.concat([existing, new_record], ignore_index=True)
            save_media_prep(updated)
            st.success(f"✅ Batch added successfully! Lot No.: **{lot_no}**")
            st.balloons()
        else:
            st.error("❌ Please fill in all required fields (Media Type and Quantity)")

# =====================================================
# TAB 2: Media Inventory
# =====================================================
with tab2:
    st.subheader("📋 Current Media Inventory")
    
    df = load_media_prep()
    
    if not df.empty:
        # Filter options
        col1, col2, col3 = st.columns(3)
        with col1:
            media_filter = st.multiselect("Filter by Media Type", options=df["Media Type"].unique().tolist())
        with col2:
            show_expired = st.checkbox("Show Expired Batches", value=False)
        with col3:
            show_zero_stock = st.checkbox("Show Zero Stock", value=False)
        
        filtered_df = df.copy()
        if media_filter:
            filtered_df = filtered_df[filtered_df["Media Type"].isin(media_filter)]
        if not show_expired:
            today = pd.Timestamp(date.today())
            filtered_df = filtered_df[filtered_df["Expiry Date"].isna() | (filtered_df["Expiry Date"] >= today)]
        if not show_zero_stock:
            filtered_df = filtered_df[filtered_df["Volume Remaining (mL)"] > 0]
        
        # Display inventory
        display_df = filtered_df.copy()
        display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m-%d") if not display_df.empty else display_df
        display_df["Expiry Date"] = display_df["Expiry Date"].dt.strftime("%Y-%m-%d") if not display_df.empty else display_df
        
        st.dataframe(display_df, use_container_width=True)
        
        # Summary stats
        st.markdown("---")
        st.subheader("📊 Inventory Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_batches = len(df)
            st.metric("Total Batches", total_batches)
        with col2:
            total_remaining = df["Volume Remaining (mL)"].sum() if not df.empty else 0
            st.metric("Total Remaining Volume", f"{total_remaining:.0f} mL")
        with col3:
            consumed = df["Volume Consumed (mL)"].sum() if not df.empty else 0
            st.metric("Total Consumed Volume", f"{consumed:.0f} mL")
        with col4:
            expired = len(df[df["Expiry Date"] < pd.Timestamp(date.today())]) if not df.empty else 0
            st.metric("Expired Batches", expired)
        
        # Download option
        st.markdown("---")
        output = io.BytesIO()
        export_df = df.copy()
        export_df["Date"] = export_df["Date"].dt.strftime("%Y-%m-%d") if not export_df.empty else export_df
        export_df["Expiry Date"] = export_df["Expiry Date"].dt.strftime("%Y-%m-%d") if not export_df.empty else export_df
        export_df.to_excel(output, index=False, engine="openpyxl")
        st.download_button("📥 Download Inventory as Excel", data=output.getvalue(), 
                          file_name="media_inventory.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No media preparation records found. Go to 'Prepare Media' tab to add batches.")

# =====================================================
# TAB 3: Consumption Log
# =====================================================
with tab3:
    st.subheader("📊 Media Consumption Log")
    
    df = load_media_prep()
    
    if not df.empty:
        # Show only batches that have been consumed
        consumed_df = df[df["Volume Consumed (mL)"] > 0].copy()
        
        if not consumed_df.empty:
            consumed_df["Consumption Percentage"] = (consumed_df["Volume Consumed (mL)"] / consumed_df["Quantity (mL)"]) * 100
            consumed_df["Date"] = consumed_df["Date"].dt.strftime("%Y-%m-%d")
            
            st.dataframe(consumed_df[["Lot No.", "Media Type", "Quantity (mL)", "Volume Consumed (mL)", 
                                      "Consumption Percentage", "Prepared By", "Date"]], use_container_width=True)
            
            # Chart
            st.markdown("---")
            st.subheader("Consumption by Media Type")
            consumption_by_type = consumed_df.groupby("Media Type")["Volume Consumed (mL)"].sum().sort_values(ascending=False)
            
            if not consumption_by_type.empty:
                import matplotlib.pyplot as plt
                fig, ax = plt.subplots(figsize=(10, 5))
                consumption_by_type.plot(kind="bar", ax=ax, color="skyblue")
                ax.set_xlabel("Media Type")
                ax.set_ylabel("Volume Consumed (mL)")
                ax.set_title("Total Media Consumption by Type")
                ax.tick_params(axis="x", rotation=45)
                st.pyplot(fig)
        else:
            st.info("No consumption records yet. Consumption happens when you perform tests in the main app.")
    else:
        st.info("No data available.")

# =====================================================
# TAB 4: Expiry Alerts
# =====================================================
with tab4:
    st.subheader("⚠️ Expiring and Expired Batches")
    
    df = load_media_prep()
    
    if not df.empty:
        today = pd.Timestamp(date.today())
        thirty_days = today + timedelta(days=30)
        
        # Expired batches
        expired = df[df["Expiry Date"] < today]
        # Expiring soon (within 30 days)
        expiring_soon = df[(df["Expiry Date"] >= today) & (df["Expiry Date"] <= thirty_days)]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### ❌ Expired Batches")
            if not expired.empty:
                expired_display = expired.copy()
                expired_display["Expiry Date"] = expired_display["Expiry Date"].dt.strftime("%Y-%m-%d")
                st.dataframe(expired_display[["Lot No.", "Media Type", "Quantity (mL)", "Volume Remaining (mL)", "Expiry Date"]], 
                            use_container_width=True)
                
                # Option to discard expired batches
                if st.button("🗑️ Delete All Expired Batches"):
                    non_expired = df[df["Expiry Date"] >= today]
                    save_media_prep(non_expired)
                    st.success("✅ Expired batches removed!")
                    st.rerun()
            else:
                st.success("No expired batches found!")
        
        with col2:
            st.markdown("### ⏰ Expiring Within 30 Days")
            if not expiring_soon.empty:
                expiring_display = expiring_soon.copy()
                expiring_display["Expiry Date"] = expiring_display["Expiry Date"].dt.strftime("%Y-%m-%d")
                st.dataframe(expiring_display[["Lot No.", "Media Type", "Volume Remaining (mL)", "Expiry Date"]], 
                            use_container_width=True)
            else:
                st.info("No batches expiring within 30 days.")
        
        # Summary
        st.markdown("---")
        total_expired_volume = expired["Volume Remaining (mL)"].sum() if not expired.empty else 0
        total_expiring_volume = expiring_soon["Volume Remaining (mL)"].sum() if not expiring_soon.empty else 0
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Wasted Volume (Expired)", f"{total_expired_volume:.0f} mL", delta="⚠️ Waste")
        with col2:
            st.metric("Volume to Use Soon", f"{total_expiring_volume:.0f} mL", delta="⚠️ Use before expiry")
    else:
        st.info("No media preparation records found.")

# =====================================================
# Sidebar Information
# =====================================================
st.sidebar.title("🏪 Microbiology Store")
st.sidebar.markdown("---")
st.sidebar.info(
    """
    **How this integrates with the main app:**
    
    1. Prepare media batches here
    2. Batches appear in main app's **Perform Test** section
    3. When you run tests, consumption is deducted automatically
    4. Track usage and expiry in this store
    
    **File:** `media_preparation.xlsx` is shared between both apps
    """
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.rerun()