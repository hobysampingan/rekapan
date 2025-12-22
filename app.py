import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(
    page_title="Order Parser - SKU Summarizer",
    page_icon="📦",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 1rem;
    }
    .platform-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 1rem;
        font-weight: 600;
        font-size: 0.875rem;
        margin-right: 0.5rem;
    }
    .badge-tiktok { background: #25F4EE; color: #000; }
    .badge-shopee { background: #EE4D2D; color: #fff; }
    .stDataFrame { border-radius: 0.5rem; }
</style>
""", unsafe_allow_html=True)

# Set the title of the Streamlit app
st.markdown('<h1 class="main-header">📦 Order Parser - SKU Quantity Summarizer</h1>', unsafe_allow_html=True)

# Supported platforms info
st.markdown("""
<div style="text-align: center; margin-bottom: 2rem;">
    <span class="platform-badge badge-tiktok">🎵 TikTok Shop</span>
    <span class="platform-badge badge-shopee">🛒 Shopee</span>
</div>
""", unsafe_allow_html=True)


def detect_platform(df):
    """Detect if the file is from TikTok or Shopee based on column names"""
    columns = df.columns.tolist()
    columns_str = ' '.join([str(col) for col in columns])
    
    # Check for Shopee specific columns (Indonesian)
    shopee_indicators = ['No. Pesanan', 'Nomor Referensi SKU', 'Status Pesanan', 'Jumlah']
    if any(indicator in columns for indicator in shopee_indicators):
        return 'shopee'
    
    # Check for TikTok specific columns
    tiktok_indicators = ['Seller SKU', 'Order ID', 'Seller sku input by the seller', 'SKU sold quantity']
    if any(indicator in columns_str for indicator in tiktok_indicators):
        return 'tiktok'
    
    return None


def get_column_mapping(platform, columns):
    """Get the correct column names for SKU and Quantity based on platform"""
    if platform == 'shopee':
        return {
            'sku': 'Nomor Referensi SKU',
            'quantity': 'Jumlah'
        }
    elif platform == 'tiktok':
        # TikTok might have different column names based on export
        # Check for the long description-style column names (when skiprows is used)
        for col in columns:
            if 'Seller sku input by the seller' in str(col):
                sku_col = col
            elif 'SKU sold quantity' in str(col):
                qty_col = col
        
        # Check for simple column names
        if 'Seller SKU' in columns:
            return {
                'sku': 'Seller SKU',
                'quantity': 'Quantity'
            }
        else:
            # Return the long column names (after skiprows)
            return {
                'sku': 'Seller sku input by the seller in the product system.',
                'quantity': 'SKU sold quantity in the order.'
            }
    
    return None


def process_file(df, platform, col_mapping):
    """Process the dataframe and return summarized SKU quantities"""
    sku_col = col_mapping['sku']
    qty_col = col_mapping['quantity']
    
    # Check if required columns exist
    if sku_col not in df.columns:
        return None, f"SKU column '{sku_col}' not found in the file."
    if qty_col not in df.columns:
        return None, f"Quantity column '{qty_col}' not found in the file."
    
    # Convert quantity column to numeric, coercing errors to NaN, then fill with 0
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0).astype(int)
    
    # Filter out empty SKUs
    df_filtered = df[[sku_col, qty_col]].copy()
    df_filtered = df_filtered[df_filtered[sku_col].notna() & (df_filtered[sku_col] != '')]
    
    # Group by SKU and sum quantity
    grouped = df_filtered.groupby(sku_col)[qty_col].sum().reset_index()
    
    # Sort alphabetically by SKU
    grouped = grouped.sort_values(by=sku_col)
    
    # Rename columns for display
    grouped = grouped.rename(columns={sku_col: 'Seller SKU', qty_col: 'Quantity'})
    
    return grouped, None


# File upload interface
uploaded_file = st.file_uploader(
    "Upload Excel file dari TikTok Shop atau Shopee (.xlsx atau .xls)", 
    type=["xlsx", "xls"],
    help="Supported: order_tiktok.xlsx, order_shopee.xlsx"
)

if uploaded_file is not None:
    try:
        # First, try reading without skiprows to detect platform
        df_detect = pd.read_excel(uploaded_file)
        platform = detect_platform(df_detect)
        
        if platform is None:
            st.error("❌ Format file tidak dikenali. Pastikan file berasal dari TikTok Shop atau Shopee.")
        else:
            # Reset file pointer
            uploaded_file.seek(0)
            
            # For TikTok, we need to skip the description row
            if platform == 'tiktok':
                df = pd.read_excel(uploaded_file, skiprows=1, header=0)
                platform_display = "🎵 TikTok Shop"
                platform_color = "#25F4EE"
            else:
                df = pd.read_excel(uploaded_file)
                platform_display = "🛒 Shopee"
                platform_color = "#EE4D2D"
            
            # Show detected platform
            st.success(f"✅ Platform terdeteksi: **{platform_display}**")
            
            # Get column mapping
            col_mapping = get_column_mapping(platform, df.columns.tolist())
            
            if col_mapping is None:
                st.error("❌ Tidak dapat menemukan kolom SKU dan Quantity yang sesuai.")
            else:
                # Process the file
                grouped, error = process_file(df, platform, col_mapping)
                
                if error:
                    st.error(f"❌ {error}")
                else:
                    # Show summary stats
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("📦 Total SKU Unik", len(grouped))
                    with col2:
                        st.metric("📊 Total Quantity", int(grouped['Quantity'].sum()))
                    with col3:
                        st.metric("📋 Total Baris Data", len(df))
                    
                    st.divider()
                    
                    # Display editable table
                    st.subheader("📋 Summarized SKU Quantities")
                    edited_grouped = st.data_editor(
                        grouped,
                        column_config={
                            "Seller SKU": st.column_config.TextColumn(disabled=True),
                            "Quantity": st.column_config.NumberColumn(min_value=0, step=1)
                        },
                        hide_index=True,
                        num_rows="fixed",
                        use_container_width=True
                    )

                    # Create HTML table for printing using edited data
                    html_table = """
                    <table style="border-collapse: collapse; width: 100%;">
                        <tr>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Seller SKU</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: right;">Qty</th>
                        </tr>
                    """

                    # Add rows to HTML table using edited data
                    for _, row in edited_grouped.iterrows():
                        html_table += f"""
                        <tr>
                            <td style="border: 1px solid black; padding: 8px;">{row['Seller SKU']}</td>
                            <td style="border: 1px solid black; padding: 8px; text-align: right;">{int(row['Quantity'])}</td>
                        </tr>
                        """

                    html_table += "</table>"

                    # Print button using HTML and JavaScript to open in new tab
                    total_qty = int(edited_grouped['Quantity'].sum())
                    total_sku = len(edited_grouped)
                    
                    print_button_html = f"""
                    <button onclick="printTable()" style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">🖨️ Print Table</button>
                    <script>
                        function printTable() {{
                            var htmlContent = `<html>
                                <head>
                                    <title>Rekap SKU - {platform_display}</title>
                                    <style>
                                        body {{ font-family: Arial, sans-serif; margin: 20px; }}
                                        h2 {{ text-align: center; margin-bottom: 0.5rem; }}
                                        .subtitle {{ text-align: center; color: #666; margin-bottom: 1rem; }}
                                        .stats {{ text-align: center; margin-bottom: 1rem; background: #f5f5f5; padding: 10px; border-radius: 8px; }}
                                        table {{ border-collapse: collapse; width: 100%; margin: 0 auto; }}
                                        th, td {{ border: 1px solid #ddd; padding: 8px; }}
                                        th {{ background-color: #f2f2f2; text-align: left; }}
                                        th:last-child {{ text-align: right; }}
                                        td:last-child {{ text-align: right; }}
                                        tr:nth-child(even) {{ background-color: #f9f9f9; }}
                                    </style>
                                </head>
                                <body>
                                    <h2>Summarized SKU Quantities</h2>
                                    <p class="subtitle">{platform_display}</p>
                                    <div class="stats">
                                        <strong>Total SKU:</strong> {total_sku} | <strong>Total Qty:</strong> {total_qty}
                                    </div>
                                    {html_table}
                                </body>
                            </html>`;
                            var newWin = window.open('', '_blank');
                            newWin.document.write(htmlContent);
                            newWin.document.close();
                            newWin.print();
                        }}
                    </script>
                    """
                    st.components.v1.html(print_button_html, height=60)
                    
                    # Download as Excel option
                    st.divider()
                    
                    # Create Excel download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        edited_grouped.to_excel(writer, index=False, sheet_name='SKU Summary')
                    
                    st.download_button(
                        label="📥 Download as Excel",
                        data=output.getvalue(),
                        file_name="sku_summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"❌ Terjadi kesalahan saat memproses file: {str(e)}")
else:
    st.info("📤 Silakan upload file Excel untuk melanjutkan.")
    
    # Show supported formats
    with st.expander("ℹ️ Format yang didukung"):
        st.markdown("""
        ### TikTok Shop
        - File export order dari TikTok Seller Center
        - Kolom yang digunakan: `Seller SKU`, `Quantity`
        
        ### Shopee
        - File export order dari Shopee Seller Centre  
        - Kolom yang digunakan: `Nomor Referensi SKU`, `Jumlah`
        """)
