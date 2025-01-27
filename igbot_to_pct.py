import re
import pandas as pd
import streamlit as st
from io import BytesIO

# Title for Streamlit app
st.title("iGBot output to PLD Files")

# Upload input files
input_file = st.file_uploader("Upload the iGBot Result file", type=["xlsx"])
file2 = st.file_uploader("Upload the POID matching file (Roaming_SC_Completion_v1.xlsx)", type=["xlsx"])
file3 = st.file_uploader("Upload the Prodef DMP file", type=["xlsx"])

if input_file and file2 and file3:
    try:
        # Step 1: Extract POID from input file name
        input_file_name = input_file.name
        try:
            extracted_poid = input_file_name.split("-")[3].split("(")[0].strip()
        except IndexError:
            st.error("Invalid input file name format. Unable to extract POID.")
            st.stop()

        # Step 2: Load file2 to match POID and get "PO Name" and "Master Keyword"
        poid_df = pd.read_excel(file2, engine="openpyxl", sheet_name="Sheet1")  # Specify engine and sheet name
        required_columns = {"POID", "POName", "Keyword"}
        if not required_columns.issubset(poid_df.columns):
            st.error(f"File2 is missing one or more required columns: {required_columns}")
            st.stop()

        # Match POID
        matched_row = poid_df[poid_df["POID"] == extracted_poid]
        if matched_row.empty:
            st.error(f"No matching POID found for '{extracted_poid}' in file2.")
            st.stop()

        final_poid = matched_row["POID"].iloc[0]
        po_name = matched_row["POName"].iloc[0]
        master_keyword = matched_row["Keyword"].iloc[0]

        # Step 3: Get ID from user
        ID = st.text_input("Enter the ID:")
        if not ID:
            st.warning("Please enter an ID to proceed.")
            st.stop()

        # Construct output file name
        output_file_name = f"PLD_{ID}_{final_poid}.xlsx"

        # Process input file
        file1 = pd.ExcelFile(input_file)

        # Create an ExcelWriter object using BytesIO for Streamlit
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter")

        # Create DataFrame for the "PO" sheet with predefined and matched values
        po_df = pd.DataFrame(
            {
                "PO ID": [final_poid],  # Matched POID from file2
                "PO Name": [po_name],  # Retrieved from file2
                "Master Keyword": [master_keyword],  # Retrieved from file2
                "Family": ["roamingSingleCountry"],  # Predefined value
                "PO Type": ["ADDON"],  # Predefined value
                "Product Category": ["b2cMobile"],  # Predefined value
                "Payment Type": ["Prepaid,Postpaid"],  # Predefined value
                "Action": ["NO_CHANGE"],  # Predefined value
            }
        )

        # Write to the Excel sheet
        po_df.to_excel(writer, sheet_name="PO", index=False)

        # Rules-Keyword
        df = pd.read_excel(file1, sheet_name="Rules-Keyword")

        # Ensure the "Short Code" column exists and manipulate it as needed
        if "Short Code" in df.columns:
            # Convert to string and strip whitespace, replace NaN with empty strings
            df["Short Code"] = df["Short Code"].astype(str).str.strip().replace("nan", "")
        else:
            # If "Short Code" column is missing, create it with default empty strings
            df["Short Code"] = ""

        # Replace any NaN with empty strings explicitly to avoid issues
        df["Short Code"] = df["Short Code"].fillna("")

        # Add the new column "Action" with the value "INSERT" for all rows
        df["Action"] = "INSERT"

        # Save the processed DataFrame to the output Excel file
        df.to_excel(writer, sheet_name="Rules-Keyword", index=False)

        # Rules-Alias
        df = pd.read_excel(file1, sheet_name="Rules-Alias")

        # Ensure the "Short Code" column exists and manipulate it as needed
        if "Short Code" in df.columns:
            # Convert to string and strip whitespace, replace NaN with empty strings
            df["Short Code"] = df["Short Code"].astype(str).str.strip().replace("nan", "")
        else:
            # If "Short Code" column is missing, create it with default empty strings
            df["Short Code"] = ""

        # Replace any NaN with empty strings explicitly to avoid issues
        df["Short Code"] = df["Short Code"].fillna("")

        # Add the new column "Action" with the value "INSERT" for all rows
        df["Action"] = "INSERT"

        # Save the processed DataFrame to the output Excel file
        df.to_excel(writer, sheet_name="Rules-Alias", index=False)

        # Rules-Header
        df = pd.read_excel(file1, sheet_name="Rules-Header")

        # Ensure the "Short Code" column exists and manipulate it as needed
        if "Ruleset Version" in df.columns:
            # Convert to string, strip whitespace, and replace "nan" with empty strings
            df["Ruleset Version"] = df["Ruleset Version"].astype(str).str.strip().replace("nan", "")
        else:
            # If "Ruleset Version" column is missing, create it with default empty strings
            df["Ruleset Version"] = ""

        # Replace any NaN with empty strings explicitly
        df["Ruleset Version"] = df["Ruleset Version"].fillna("")

        # Convert to numeric, coercing errors to NaN, then fill NaN with 0 and convert to integer
        df["Ruleset Version"] = pd.to_numeric(df["Ruleset Version"], errors="coerce").fillna(0).astype(int)

        # Add the new column "Action" with the value "INSERT" for all rows
        df["Action"] = "INSERT"

        # Save the processed DataFrame to the output Excel file
        df.to_excel(writer, sheet_name="Rules-Header", index=False)

        # Process Rules-PCRF sheet
        df_pcrf = pd.read_excel(file1, sheet_name="Rules-PCRF")
        df_pcrf["Ruleset ShortName"] = df_pcrf["Ruleset ShortName"].astype(str).str.strip()

        # Ensure Lifetime and MaxLifetime columns exist
        columns_to_convert = ["LifeTime Validity", "MaxLife Time"]
        for col in columns_to_convert:
            if col in df_pcrf.columns:
                # Convert to string, strip whitespace, and replace NaN with empty strings
                df_pcrf[col] = df_pcrf[col].fillna("").astype(str).str.strip().replace("nan", "")
            else:
                # If column is missing, create it with default empty strings
                df_pcrf[col] = ""

        # Add the new column "Action" with the value "INSERT" for all rows
        df_pcrf["Action"] = "INSERT"

        # Save the updated DataFrame to the Excel sheet as PCRF
        df_pcrf.to_excel(writer, sheet_name="PCRF", index=False)

        # Handle specific sheets
        try:
            df = pd.read_excel(file1, sheet_name="Rules-Cases-Condition")
            if "OpIndex" in df.columns:
                df["OpIndex"] = pd.to_numeric(df["OpIndex"], errors="coerce").astype("Int64")
            # Add the new column "Action" with value "INSERT" for all rows
            df["Action"] = "INSERT"
            
            df.to_excel(writer, sheet_name="Rules-Cases-Condition", index=False)
        except Exception as e:
            st.error(f"Error processing 'Rules-Cases-Condition': {e}")

        # Rules-Cases-Success
        try:
            df = pd.read_excel(file1, sheet_name="Rules-Cases-Success")
            if "OpIndex" in df.columns:
                df["OpIndex"] = pd.to_numeric(df["OpIndex"], errors="coerce").astype("Int64")
            if "Ruleset ShortName" in df.columns:
                df["Exit Value"] = df["Ruleset ShortName"].apply(
                    lambda x: "1" if pd.notna(x) and x.strip() != "" else ""
                )
            # Add the new column "Action" with the value "INSERT" for all rows
            df["Action"] = "INSERT"

            df.to_excel(writer, sheet_name="Rules-Cases-Success", index=False)
        except Exception as e:
            st.error(f"Error processing 'Rules-Cases-Success': {e}")

        # Example sheet creation: Rules-Messages
        messages_df = pd.DataFrame(
            {
                "PO ID": ["sample"],
                "Ruleset ShortName": ["sample"],
                "Order Status": ["sample"],
                "Order Type": ["sample"],
                "Sender Address": ["sample"],
                "Channel": ["sample"],
                "Message Content Index": ["sample"],
                "Message Content": ["sample"],
                "Action": ["NO CHANGE"],
            }
        )
        messages_df.to_excel(writer, sheet_name="Rules-Messages", index=False)

        # Sheet 9: Rules-Price-Mapping
        df = pd.read_excel(file1, sheet_name="Rules-Price-Mapping")

        # Convert "Price" column to integers
        df["Price"] = (
            df["Price"]
            .astype(str)  # Convert to string to handle potential non-string values
            .str.replace(",", "")  # Remove commas
            .pipe(pd.to_numeric, errors="coerce")  # Convert to numeric, coercing errors to NaN
            .astype("Int64")  # Convert to integer (nullable type)
        )

        # Ensure the "SID" column exists and manipulate it as needed
        if "SID" in df.columns:
            # Convert to string, strip whitespace, replace "nan", and handle numeric values
            df["SID"] = (
                df["SID"]
                .astype(str)
                .str.strip()
                .replace("nan", "")
                .apply(lambda x: str(int(float(x))) if x.replace(".", "").isdigit() else x)
            )
        else:
            # If "SID" column is missing, create it with default empty strings
            df["SID"] = ""

        # Replace any NaN with empty strings explicitly
        df["SID"] = df["SID"].fillna("")

        # Add the new column "Action" with the value "INSERT" for all rows
        df["Action"] = "INSERT"

        # Save the modified DataFrame to the Excel sheet
        df.to_excel(writer, sheet_name="Rules-Price-Mapping", index=False)

        # Sheet 10: Rules-Renewal
        df = pd.read_excel(file1, sheet_name="Rules-Renewal")

        # Convert "Max Cycle" and "Period" columns to integers
        df["Max Cycle"] = pd.to_numeric(df["Max Cycle"], errors="coerce").astype("Int64")
        df["Period"] = pd.to_numeric(df["Period"], errors="coerce").astype("Int64")

        # Remove commas and decimals from "Amount" and truncate decimals
        if "Amount" in df.columns:
            # Remove commas, keep only numeric part, and truncate decimals
            df["Amount"] = (
                df["Amount"]
                .str.replace(",", "", regex=False)  # Remove commas
                .str.split(".", n=1).str[0]        # Remove decimals by splitting
            )
            # Convert to integer
            df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").astype("Int64")
        else:
            df["Amount"] = None  # Handle cases where "Amount" column is missing

        # Ensure the "Reg Subaction" column exists and manipulate it as needed
        if "Reg Subaction" in df.columns:
            # Convert to string and strip whitespace, replace NaN with empty strings
            df["Reg Subaction"] = df["Reg Subaction"].astype(str).str.strip().replace("nan", "")
        else:
            # If "Reg Subaction" column is missing, create it with default empty strings
            df["Reg Subaction"] = ""

        # Replace any NaN with empty strings explicitly to avoid issues
        df["Reg Subaction"] = df["Reg Subaction"].fillna("")

        # Add the new column "Action" with the value "INSERT" for all rows
        df["Action"] = "INSERT"

        # Save the modified DataFrame to the Excel sheet
        df.to_excel(writer, sheet_name="Rules-Renewal", index=False)

        # Sheet 11: Rules-GSI GRP Pack
        gsi_grp_pack_df = pd.DataFrame(
            {
                "Ruleset ShortName": ["sample"],  # First row value
                "GSI GRP Pack-Group ID": ["sample"],  # First row value
                "Action": ["NO_CHANGE"],  # First row value
            }
        )
        gsi_grp_pack_df.to_excel(writer, sheet_name="Rules-GSI GRP Pack", index=False)

        # Sheet 12: Rules-Location Group
        location_group_df = pd.DataFrame(
            {
                "Ruleset ShortName": ["sample"],
                "Package Group": ["sample"],
                "Microcluster ID": ["sample"],
                "Action": ["NO_CHANGE"],
            }
        )
        location_group_df.to_excel(writer, sheet_name="Rules-Location Group", index=False)

        # Sheet 13: Rebuy-Out
        rebuy_out_df = pd.DataFrame(
            {
                "Target PO ID": ["sample"],
                "Target Ruleset ShortName": ["sample"],
                "Target MPP": ["sample"],
                "Target Group": ["sample"],
                "Service Type": ["sample"],
                "Rebuy Price": ["sample"],
                "Allow Rebuy": ["sample"],
                "Rebuy Option": ["sample"],
                "Product Family": ["sample"],
                "Source PO ID": ["sample"],
                "Source Ruleset ShortName": ["sample"],
                "Source MPP": ["sample"],
                "Source Group": ["sample"],
                "Vice Versa Consent": ["sample"],
                "Action": ["NO_CHANGE"],
            }
        )
        rebuy_out_df.to_excel(writer, sheet_name="Rebuy-Out", index=False)

        # Sheet 14: Rebuy-Association
        rebuy_association_df = pd.DataFrame(
            {
                "Target PO ID": ["sample"],
                "Target Ruleset ShortName": ["sample"],
                "Target MPP": ["sample"],
                "Target Group": ["sample"],
                "Service Type": ["sample"],
                "Rebuy Price": ["sample"],
                "Allow Rebuy": ["sample"],
                "Rebuy Option": ["sample"],
                "Product Family": ["sample"],
                "Source PO ID": ["sample"],
                "Source Ruleset ShortName": ["sample"],
                "Source MPP": ["sample"],
                "Source Group": ["sample"],
                "Vice Versa Consent": ["sample"],
                "Action": ["NO_CHANGE"],
            }
        )
        rebuy_association_df.to_excel(writer, sheet_name="Rebuy-Association", index=False)

        # Sheet 15: Incompatibility
        incompatibility_df = pd.DataFrame(
            {
                "ID": ["sample"],
                "Target PO/RulesetShortName": ["sample"],
                "Source Family": ["sample"],
                "Source PO/RulesetShortName": ["sample"],
                "Action": ["NO_CHANGE"],
            }
        )
        incompatibility_df.to_excel(writer, sheet_name="Incompatibility", index=False)

        # Sheet 16: Library-Addon-Name
        df = pd.read_excel(file1, sheet_name="Rules-Library-Addon")

        # List of columns to process to maintain as string
        columns_to_process = ["Master Shortcode", "Active Period Length", "Grace Period"]

        for col in columns_to_process:
            # Ensure the column exists
            if col in df.columns:
                # Convert to string, strip whitespace, replace NaN with empty strings
                df[col] = df[col].fillna("").astype(str).str.strip().replace("nan", "")
            else:
                # If column is missing, create it with default empty strings
                df[col] = ""

        # Add the new column "Action" with the value "INSERT" for all rows
        df["Action"] = "INSERT"

        # Save the modified DataFrame to the Excel sheet
        df.to_excel(writer, sheet_name="Library-Addon-Name", index=False)

        # Sheet 17: Library-Addon-DA
        df_library_addon_da = pd.read_excel(file3, engine="openpyxl", sheet_name="Library AddOn_DA")
        df_library_addon_da["Ruleset ShortName"] = df_library_addon_da["Ruleset ShortName"].astype(str).str.strip()
        df_library_addon_da["daid"] = df_library_addon_da["daid"].astype(str)
        df_library_addon_da["Action"] = "INSERT"

        # Ensure "Ruleset ShortName" is updated using PCRF
        def update_ruleset_shortname(row, df_pcrf):
            current_shortname = row["Ruleset ShortName"]
            
            if "PRE" in current_shortname:
                # Match a "Ruleset ShortName" in PCRF containing "PRE"
                match = df_pcrf[df_pcrf["Ruleset ShortName"].str.contains("PRE")]
                if not match.empty:
                    return match.iloc[0]["Ruleset ShortName"]  # Replace with the first matching "PRE"
            elif "ACT" in current_shortname:
                # Match a "Ruleset ShortName" in PCRF containing "ACT"
                match = df_pcrf[df_pcrf["Ruleset ShortName"].str.contains("ACT")]
                if not match.empty:
                    return match.iloc[0]["Ruleset ShortName"]  # Replace with the first matching "ACT"
            else:
                # Take the first "Ruleset ShortName" from PCRF that does not contain "PRE" or "ACT"
                non_pre_act = df_pcrf[~df_pcrf["Ruleset ShortName"].str.contains("PRE|ACT")]
                if not non_pre_act.empty:
                    return non_pre_act.iloc[0]["Ruleset ShortName"]
            
            # If no match, return the current value (fallback)
            return current_shortname
        
        # Apply the update function to "Ruleset ShortName"
        df_library_addon_da["Ruleset ShortName"] = df_library_addon_da.apply(update_ruleset_shortname, axis=1, df_pcrf=df_pcrf)
        
        df_library_addon_da.to_excel(writer, sheet_name="Library-Addon-DA", index=False)

        # Sheet 18: Library-Addon-UCUT
        library_addon_ucut_df = pd.DataFrame(
            {
                "Ruleset ShortName": ["sample"],
                "PO ID": ["sample"],
                "Quota Name": ["sample"],
                "UCUT ID": ["sample"],
                "Internal Description Bahasa": ["sample"],
                "External Description Bahasa": ["sample"],
                "Internal Description English": ["sample"],
                "External Description English": ["sample"],
                "Visibility": ["sample"],
                "Custom": ["sample"],
                "Initial Value": ["sample"],
                "Unlimited Benefit Flag": ["sample"],
                "Action": ["NO_CHANGE"],
            }
        )
        library_addon_ucut_df.to_excel(writer, sheet_name="Library-Addon-UCUT", index=False)

        # Sheet 19: Standalone - copy from file3.xlsx "StandAlone"
        standalone_df= pd.read_excel(file3, engine="openpyxl", sheet_name="StandAlone")
        standalone_df["Action"] = "INSERT"  # Add "Action" column with value "INSERT"

        # Convert 'Value', 'UOM', 'Validity' column to string in standalone_df
        standalone_df['Value'] = standalone_df['Value'].astype(str)
        standalone_df['UOM'] = standalone_df['UOM'].astype(str)
        standalone_df['Validity'] = standalone_df['Validity'].astype(str)

        standalone_df.to_excel(writer, sheet_name="Standalone", index=False)

        # Sheet 20: Blacklist-Gift-Promocodes
        blacklist_gift_promocodes_df = pd.DataFrame(
            [{"Ruleset ShortName": "sample", "Coherence Key": "sample", "Promo Codes": "sample", "Action": "NO_CHANGE"}]
        )
        blacklist_gift_promocodes_df.to_excel(writer, sheet_name="Blacklist-Gift-Promocodes", index=False)

        # Sheet 21: Blacklist-Promocodes
        blacklist_promocodes_df = pd.DataFrame(
            [{"PO ID": "sample", "Command/Keyword": "sample", "Promo Codes": "sample", "Action": "NO_CHANGE"}]
        )
        blacklist_promocodes_df.to_excel(writer, sheet_name="Blacklist-Promocodes", index=False)

        # Sheet 22: MYIM3-UNREG
        myim3_unreg_df = pd.DataFrame(
            [
                {
                    "Ruleset ShortName": "sample",
                    "Keyword": "sample",
                    "Shortcode": "sample",
                    "Unreg Flag": "sample",
                    "Buy Extra Flag": "sample",
                    "Action": "NO_CHANGE",
                }
            ]
        )
        myim3_unreg_df.to_excel(writer, sheet_name="MYIM3-UNREG", index=False)

        # Sheet 23: ExtraPOConfig
        extrapoconfig_df = pd.DataFrame(
            [{"Ruleset ShortName": "sample", "Extra PO Keyword": "sample", "Action": "NO_CHANGE"}]
        )
        extrapoconfig_df.to_excel(writer, sheet_name="ExtraPOConfig", index=False)

        # Sheet 24: Keyword-Global-Variable
        keyword_global_variable_df = pd.DataFrame(
            [
                {
                    "PO ID": "sample",
                    "Keyword": "sample",
                    "Global Variable Type": "sample",
                    "Value": "sample",
                    "Keyword Type": "sample",
                    "Action": "NO_CHANGE",
                }
            ]
        )
        keyword_global_variable_df.to_excel(writer, sheet_name="Keyword-Global-Variable", index=False)

        # Sheet 25: UMB-Push-Category
        umb_push_category_df = pd.DataFrame(
            [
                {
                    "Ruleset ShortName": "sample",
                    "Coherence Key": "sample",
                    "Group Category": "sample",
                    "Short Code": "sample",
                    "Show Unit": "sample",
                    "Action": "NO_CHANGE",
                }
            ]
        )
        umb_push_category_df.to_excel(writer, sheet_name="UMB-Push-Category", index=False)

        # Sheet 26: Avatar-Channel
        avatar_channel_df = pd.DataFrame(
            [
                {
                    "PO ID": "sample",
                    "Ruleset ShortName": "sample",
                    "Keyword": "sample",
                    "Commercial Name": "sample",
                    "Short Code": "sample",
                    "PVR ID": "sample",
                    "Price": "sample",
                    "Action": "NO_CHANGE",
                }
            ]
        )
        avatar_channel_df.to_excel(writer, sheet_name="Avatar-Channel", index=False)

        # Sheet 27: Dormant-Config
        dormant_config_df = pd.DataFrame(
            [{"Ruleset ShortName": "sample", "Keyword": "sample", "Short Code": "sample", "Pvr": "sample", "Action": "NO_CHANGE"}]
        )
        dormant_config_df.to_excel(writer, sheet_name="Dormant-Config", index=False)


        writer.close()
        st.download_button(
            label="Download Processed Excel",
            data=output.getvalue(),
            file_name=output_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")
