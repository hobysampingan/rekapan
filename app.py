import streamlit as st
import pandas as pd
import io

# Set the title of the Streamlit app
st.title("Excel Order Parser - SKU Quantity Summarizer")

# File upload interface
uploaded_file = st.file_uploader("Upload an Excel file (.xlsx or .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file, skipping the description row and using the header row
        # skiprows=1 means skip row 1 (description), header=0 means use row 0 as header
        df = pd.read_excel(uploaded_file, skiprows=1, header=0)

        # Define the actual column names based on the file structure
        seller_sku_col = 'Seller sku input by the seller in the product system.'
        quantity_col = 'SKU sold quantity in the order.'

        # Check if required columns exist
        if seller_sku_col not in df.columns or quantity_col not in df.columns:
            st.error("Error: Required columns 'Seller SKU' and 'Quantity' not found in the uploaded file.")
        else:
            # Convert quantity column to numeric, coercing errors to NaN, then fill with 0
            df[quantity_col] = pd.to_numeric(df[quantity_col], errors="coerce").fillna(0).astype(int)

            # Filter to only the required columns
            df_filtered = df[[seller_sku_col, quantity_col]]

            # Group by seller SKU and sum quantity
            grouped = df_filtered.groupby(seller_sku_col)[quantity_col].sum().reset_index()

            # Sort alphabetically by seller SKU
            grouped = grouped.sort_values(by=seller_sku_col)

            # Rename columns for display
            grouped = grouped.rename(columns={seller_sku_col: 'Seller SKU', quantity_col: 'Quantity'})

            # Display editable table
            st.subheader("Summarized SKU Quantities")
            edited_grouped = st.data_editor(
                grouped,
                column_config={
                    "Seller SKU": st.column_config.TextColumn(disabled=True),
                    "Quantity": st.column_config.NumberColumn(min_value=0, step=1)
                },
                hide_index=True,
                num_rows="fixed"
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
            print_button_html = f"""
            <button onclick="printTable()" style="background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">Print Table</button>
            <script>
                function printTable() {{
                    var htmlContent = `<html>
                        <head>
                            <title>Print Table</title>
                            <style>
                                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                                h2 {{ text-align: center; }}
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
            st.components.v1.html(print_button_html)

    except Exception as e:
        st.error(f"An error occurred while processing the file: {str(e)}")
else:
    st.info("Please upload an Excel file to proceed.")
