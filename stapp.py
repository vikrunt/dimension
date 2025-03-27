# # # import streamlit as st
# # # import pandas as pd
# # # import math
# # # import io
# # #
# # #
# # # # Functions for calculations
# # # def calculate_opr(value, min_value, max_value):
# # #     if value == 0:
# # #         return 0
# # #     if value <= min_value:
# # #         return 1
# # #     opr = math.ceil(value / min_value)
# # #     if opr > 1 and value / (opr - 1) < max_value:
# # #         return opr - 1
# # #     return opr
# # #
# # #
# # # def calculate_tab(max_opr, master_tab=1):
# # #     if max_opr == 0:
# # #         return 0
# # #     if 8 <= max_opr < 16:
# # #         return max_opr + 2 - master_tab
# # #     elif max_opr > 15:
# # #         return max_opr + 3 - master_tab
# # #     else:
# # #         return max_opr + 1 - master_tab
# # #
# # #
# # # def process_range(df, min_value, max_value, range_label, day_headers):
# # #     for day, headers in day_headers.items():
# # #         for header in headers:
# # #             opr_col = f"{header} Opr ({range_label})"
# # #             df[opr_col] = df[header].apply(lambda x: calculate_opr(x, min_value, max_value))
# # #
# # #     opr_columns = [f"{header} Opr ({range_label})" for day, headers in day_headers.items() for header in headers]
# # #     df[f"All-Day Max Opr ({range_label})"] = df[opr_columns].max(axis=1)
# # #
# # #     df[f"Tab ({range_label})"] = df[f"All-Day Max Opr ({range_label})"].apply(lambda x: calculate_tab(x))
# # #     df[f"Master Tab ({range_label})"] = 1
# # #     df[f"SIM ({range_label})"] = 1
# # #     df[f"IRIS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
# # #     df[f"FPS ({range_label})"] = 1
# # #     df[f"OTG ({range_label})"] = df[f"IRIS ({range_label})"] * 2
# # #     df[f"Hologram ({range_label})"] = df["Total Candidate"].apply(lambda x: math.ceil(x / 100) + 1)
# # #     df[f"Id Card ({range_label})"] = df[f"All-Day Max Opr ({range_label})"] + 1
# # #     df[f"Jacket ({range_label})"] = df[f"All-Day Max Opr ({range_label})"]
# # #     df["Range"] = range_label
# # #     return df
# # #
# # #
# # # # Streamlit App
# # # st.title("Operational Ratio Processing")
# # #
# # # # File upload
# # # uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
# # #
# # # if uploaded_file:
# # #     # Load the file into a DataFrame
# # #     df = pd.read_excel(uploaded_file)
# # #     df.fillna(0, inplace=True)
# # #
# # #     # Define headers for each day
# # #     day_headers = {
# # #         "Day-1": ["19 Jan'25","20 Jan'25"],
# # #
# # #     }
# # #
# # #     # Calculate "Total Candidate"
# # #     df["Total Candidate"] = df[[header for headers in day_headers.values() for header in headers]].sum(axis=1)
# # #
# # #     # Define ranges
# # #     ranges = [(61, 61), (66, 66), (71, 71), (76, 76), (81, 81), (86, 86), (91, 91), (96, 96), (101, 101)]
# # #     detailed_results = []
# # #     grand_totals_combined = []
# # #
# # #     # Process each range
# # #     for min_val, max_val in ranges:
# # #         label = f"{min_val}-{max_val}"
# # #         processed_df = process_range(df.copy(), min_val, max_val, label, day_headers)
# # #         detailed_results.append((label, processed_df))
# # #
# # #         totals = {
# # #             "Range": label,
# # #             **{f"{header} Opr": processed_df[f"{header} Opr ({label})"].sum() for day, headers in day_headers.items()
# # #                for header in headers},
# # #             "All-Day Max Opr": processed_df[f"All-Day Max Opr ({label})"].sum(),
# # #             "Tab": processed_df[f"Tab ({label})"].sum(),
# # #             "Master Tab": processed_df[f"Master Tab ({label})"].sum(),
# # #             "SIM": processed_df[f"SIM ({label})"].sum(),
# # #             "IRIS": processed_df[f"IRIS ({label})"].sum(),
# # #             "FPS": processed_df[f"FPS ({label})"].sum(),
# # #             "OTG": processed_df[f"OTG ({label})"].sum(),
# # #             "Hologram": processed_df[f"Hologram ({label})"].sum(),
# # #             "Id Card": processed_df[f"Id Card ({label})"].sum(),
# # #             "Jacket": processed_df[f"Jacket ({label})"].sum(),
# # #         }
# # #         grand_totals_combined.append(totals)
# # #
# # #     # Create Grand Totals DataFrame
# # #     grand_totals_final = pd.DataFrame(grand_totals_combined)
# # #
# # #     # Save to Excel in-memory
# # #     output = io.BytesIO()
# # #     with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
# # #         for label, detailed_df in detailed_results:
# # #             detailed_df.to_excel(writer, sheet_name=label, index=False)
# # #         grand_totals_final.to_excel(writer, sheet_name="Grand Totals", index=False)
# # #
# # #     # Button to download the result
# # #     st.download_button(
# # #         label="Download Processed Excel File",
# # #         data=output.getvalue(),
# # #         file_name="processed_data.xlsx",
# # #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
# # #     )
# #
# #
# # import streamlit as st
# # import pandas as pd
# # import math
# # import io
# #
# #
# # # Functions for calculations
# # def calculate_opr(value, min_value, max_value):
# #     if value == 0:
# #         return 0
# #     if value <= min_value:
# #         return 1
# #     opr = math.ceil(value / min_value)
# #     if opr > 1 and value / (opr - 1) < max_value:
# #         return opr - 1
# #     return opr
# #
# #
# # def calculate_tab(max_opr, master_tab=1):
# #     if max_opr == 0:
# #         return 0
# #     if 8 <= max_opr < 16:
# #         return max_opr + 2 - master_tab
# #     elif max_opr > 15:
# #         return max_opr + 3 - master_tab
# #     else:
# #         return max_opr + 1 - master_tab
# #
# #
# # def process_range(df, min_value, max_value, range_label, day_headers, calculation_mode):
# #     for day, headers in day_headers.items():
# #         for header in headers:
# #             opr_col = f"{header} Opr ({range_label})"
# #             df[opr_col] = df[header].apply(lambda x: calculate_opr(x, min_value, max_value))
# #
# #     opr_columns = [f"{header} Opr ({range_label})" for day, headers in day_headers.items() for header in headers]
# #     df[f"All-Day Max Opr ({range_label})"] = df[opr_columns].max(axis=1)
# #
# #     df[f"Tab ({range_label})"] = df[f"All-Day Max Opr ({range_label})"].apply(lambda x: calculate_tab(x))
# #     df[f"Master Tab ({range_label})"] = 1
# #     df[f"SIM ({range_label})"] = 1
# #
# #     if calculation_mode == "IRIS":
# #         df[f"IRIS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
# #         df[f"FPS ({range_label})"] = 1
# #     elif calculation_mode == "FPS":
# #         df[f"FPS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
# #
# #     df[f"OTG ({range_label})"] = df[f"{calculation_mode} ({range_label})"] * 2
# #     df[f"Hologram ({range_label})"] = df["Total Candidate"].apply(lambda x: math.ceil(x / 100) + 1)
# #     df[f"Id Card ({range_label})"] = df[f"All-Day Max Opr ({range_label})"] + 1
# #     df[f"Jacket ({range_label})"] = df[f"All-Day Max Opr ({range_label})"]
# #     df["Range"] = range_label
# #     return df
# #
# #
# # # Streamlit App
# # st.title("Operational Ratio Processing")
# #
# # # File upload
# # uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
# #
# # if uploaded_file:
# #     # Load the file into a DataFrame
# #     df = pd.read_excel(uploaded_file)
# #     df.fillna(0, inplace=True)
# #
# #     # Allow the user to select columns for day headers
# #     st.header("Select Columns for Day Headers")
# #     day_headers = {}
# #     for day in ["Day-1"]:  # You can add more days if needed
# #         selected_cols = st.multiselect(f"Select columns for {day}", df.columns.tolist())
# #         if selected_cols:
# #             day_headers[day] = selected_cols
# #
# #     # Ensure day_headers is not empty
# #     if not day_headers:
# #         st.warning("Please select at least one column for day headers.")
# #         st.stop()
# #
# #     # Allow the user to choose between IRIS and FPS
# #     calculation_mode = st.radio("Choose Calculation Mode", options=["IRIS", "FPS"], index=0)
# #
# #     # Calculate "Total Candidate"
# #     df["Total Candidate"] = df[[header for headers in day_headers.values() for header in headers]].sum(axis=1)
# #
# #     # Define ranges
# #     ranges = [(61, 61), (66, 66), (71, 71), (76, 76), (81, 81), (86, 86), (91, 91), (96, 96), (101, 101)]
# #     detailed_results = []
# #     grand_totals_combined = []
# #
# #     # Process each range
# #     for min_val, max_val in ranges:
# #         label = f"{min_val}-{max_val}"
# #         processed_df = process_range(df.copy(), min_val, max_val, label, day_headers, calculation_mode)
# #         detailed_results.append((label, processed_df))
# #
# #         totals = {
# #             "Range": label,
# #             **{f"{header} Opr": processed_df[f"{header} Opr ({label})"].sum() for day, headers in day_headers.items()
# #                for header in headers},
# #             "All-Day Max Opr": processed_df[f"All-Day Max Opr ({label})"].sum(),
# #             "Tab": processed_df[f"Tab ({label})"].sum(),
# #             "Master Tab": processed_df[f"Master Tab ({label})"].sum(),
# #             "SIM": processed_df[f"SIM ({label})"].sum(),
# #             f"{calculation_mode}": processed_df[f"{calculation_mode} ({label})"].sum(),
# #             "OTG": processed_df[f"OTG ({label})"].sum(),
# #             "Hologram": processed_df[f"Hologram ({label})"].sum(),
# #             "Id Card": processed_df[f"Id Card ({label})"].sum(),
# #             "Jacket": processed_df[f"Jacket ({label})"].sum(),
# #         }
# #         grand_totals_combined.append(totals)
# #
# #     # Create Grand Totals DataFrame
# #     grand_totals_final = pd.DataFrame(grand_totals_combined)
# #
# #     # Save to Excel in-memory
# #     output = io.BytesIO()
# #     with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
# #         for label, detailed_df in detailed_results:
# #             if calculation_mode == "FPS":
# #                 detailed_df = detailed_df.drop(columns=[col for col in detailed_df.columns if "IRIS" in col])
# #             detailed_df.to_excel(writer, sheet_name=label, index=False)
# #         grand_totals_final.to_excel(writer, sheet_name="Grand Totals", index=False)
# #
# #     # Button to download the result
# #     st.download_button(
# #         label="Download Processed Excel File",
# #         data=output.getvalue(),
# #         file_name="processed_data.xlsx",
# #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
# #     )
#
#

















#
# import streamlit as st
# import pandas as pd
# import math
# import io
#
#
# # Functions for calculations
# def calculate_opr(value, min_value, max_value):
#     if value == 0:
#         return 0
#     if value <= min_value:
#         return 1
#     opr = math.ceil(value / min_value)
#     if opr > 1 and value / (opr - 1) < max_value:
#         return opr - 1
#     return opr
#
#
# def calculate_tab(max_opr, master_tab=1):
#     if max_opr == 0:
#         return 0
#     if 8 <= max_opr < 16:
#         return max_opr + 2 - master_tab
#     elif max_opr > 15:
#         return max_opr + 3 - master_tab
#     else:
#         return max_opr + 1 - master_tab
#
#
# def process_range(df, min_value, max_value, range_label, day_headers, calculation_mode):
#     for day, headers in day_headers.items():
#         for header in headers:
#             opr_col = f"{header} Opr ({range_label})"
#             df[opr_col] = df[header].apply(lambda x: calculate_opr(x, min_value, max_value))
#
#     opr_columns = [f"{header} Opr ({range_label})" for day, headers in day_headers.items() for header in headers]
#     df[f"All-Day Max Opr ({range_label})"] = df[opr_columns].max(axis=1)
#
#     df[f"Tab ({range_label})"] = df[f"All-Day Max Opr ({range_label})"].apply(lambda x: calculate_tab(x))
#     df[f"Master Tab ({range_label})"] = 1
#     df[f"SIM ({range_label})"] = 1
#
#     if calculation_mode == "IRIS":
#         df[f"IRIS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
#         df[f"FPS ({range_label})"] = 1
#     elif calculation_mode == "FPS":
#         df[f"FPS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
#
#     df[f"OTG ({range_label})"] = df[f"{calculation_mode} ({range_label})"] * 2
#     df[f"Hologram ({range_label})"] = df["Total Candidate"].apply(lambda x: math.ceil(x / 100) + 1)
#     df[f"Id Card ({range_label})"] = df[f"All-Day Max Opr ({range_label})"] + 1
#     df[f"Jacket ({range_label})"] = df[f"All-Day Max Opr ({range_label})"]
#     df["Range"] = range_label
#     return df
#
#
# # Streamlit App
# st.title("Operational Ratio Processing")
#
# # File upload
# uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
#
# if uploaded_file:
#     # Load the file into a DataFrame
#     df = pd.read_excel(uploaded_file)
#     df.fillna(0, inplace=True)
#
#     # Allow the user to select columns for day headers
#     st.header("Select Columns for Day Headers")
#     day_headers = {}
#     for day in ["Day-1"]:  # You can add more days if needed
#         selected_cols = st.multiselect(f"Select columns for {day}", df.columns.tolist())
#         if selected_cols:
#             day_headers[day] = selected_cols
#
#     # Ensure day_headers is not empty
#     if not day_headers:
#         st.warning("Please select at least one column for day headers.")
#         st.stop()
#
#     # Allow the user to choose between IRIS and FPS
#     calculation_mode = st.radio("Choose Calculation Mode", options=["IRIS", "FPS"], index=0)
#
#     # Default and customizable ranges
#     st.header("Select Ranges")
#     default_ranges = [(61, 61), (66, 66), (71, 71), (76, 76), (81, 81), (86, 86), (91, 91), (96, 96), (101, 101)]
#     default_ranges_display = ", ".join([f"({min_val}, {max_val})" for min_val, max_val in default_ranges])
#     st.write(f"Default Ranges: {default_ranges_display}")
#
#     custom_ranges_input = st.text_input(
#         "Enter custom ranges (e.g., 50-55, 60-65). Leave blank to use default ranges."
#     )
#     if custom_ranges_input.strip():
#         try:
#             custom_ranges = [
#                 tuple(map(int, r.strip().split("-"))) for r in custom_ranges_input.split(",")
#             ]
#         except ValueError:
#             st.error("Invalid range format. Please use the format 'min-max' separated by commas.")
#             st.stop()
#     else:
#         custom_ranges = default_ranges
#
#     # Calculate "Total Candidate"
#     df["Total Candidate"] = df[[header for headers in day_headers.values() for header in headers]].sum(axis=1)
#
#     # Process ranges
#     detailed_results = []
#     grand_totals_combined = []
#
#     for min_val, max_val in custom_ranges:
#         label = f"{min_val}-{max_val}"
#         processed_df = process_range(df.copy(), min_val, max_val, label, day_headers, calculation_mode)
#         detailed_results.append((label, processed_df))
#
#         totals = {
#             "Range": label,
#             **{f"{header} Opr": processed_df[f"{header} Opr ({label})"].sum() for day, headers in day_headers.items()
#                for header in headers},
#             "All-Day Max Opr": processed_df[f"All-Day Max Opr ({label})"].sum(),
#             "Tab": processed_df[f"Tab ({label})"].sum(),
#             "Master Tab": processed_df[f"Master Tab ({label})"].sum(),
#             "SIM": processed_df[f"SIM ({label})"].sum(),
#             f"{calculation_mode}": processed_df[f"{calculation_mode} ({label})"].sum(),
#             "OTG": processed_df[f"OTG ({label})"].sum(),
#             "Hologram": processed_df[f"Hologram ({label})"].sum(),
#             "Id Card": processed_df[f"Id Card ({label})"].sum(),
#             "Jacket": processed_df[f"Jacket ({label})"].sum(),
#         }
#         grand_totals_combined.append(totals)
#
#     # Create Grand Totals DataFrame
#     grand_totals_final = pd.DataFrame(grand_totals_combined)
#
#     # Save to Excel in-memory
#     output = io.BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
#         for label, detailed_df in detailed_results:
#             if calculation_mode == "FPS":
#                 detailed_df = detailed_df.drop(columns=[col for col in detailed_df.columns if "IRIS" in col])
#             detailed_df.to_excel(writer, sheet_name=label, index=False)
#         grand_totals_final.to_excel(writer, sheet_name="Grand Totals", index=False)
#
#     # Button to download the result
#     st.download_button(
#         label="Download Processed Excel File",
#         data=output.getvalue(),
#         file_name="processed_data.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     )
#















#
# import streamlit as st
# import pandas as pd
# import math
# import io
#
#
# # Functions for calculations
# def calculate_opr(value, min_value, max_value):
#     if value == 0:
#         return 0
#     if value <= min_value:
#         return 1
#     opr = math.ceil(value / min_value)
#     if opr > 1 and value / (opr - 1) < max_value:
#         return opr - 1
#     return opr
#
#
# def calculate_tab(max_opr, master_tab=1):
#     if max_opr == 0:
#         return 0
#     if 8 <= max_opr < 16:
#         return max_opr + 2 - master_tab
#     elif max_opr > 15:
#         return max_opr + 3 - master_tab
#     else:
#         return max_opr + 1 - master_tab
#
#
# def process_range(df, min_value, max_value, range_label, day_headers, calculation_mode):
#     for day, headers in day_headers.items():
#         for header in headers:
#             opr_col = f"{header} Opr ({range_label})"
#             df[opr_col] = df[header].apply(lambda x: calculate_opr(x, min_value, max_value))
#
#     opr_columns = [f"{header} Opr ({range_label})" for day, headers in day_headers.items() for header in headers]
#     df[f"All-Day Max Opr ({range_label})"] = df[opr_columns].max(axis=1)
#
#     df[f"Tab ({range_label})"] = df[f"All-Day Max Opr ({range_label})"].apply(lambda x: calculate_tab(x))
#     df[f"Master Tab ({range_label})"] = 1
#     df[f"SIM ({range_label})"] = 1
#
#     if calculation_mode == "IRIS":
#         df[f"IRIS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
#         df[f"FPS ({range_label})"] = 1
#     elif calculation_mode == "FPS":
#         df[f"FPS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
#
#     df[f"OTG ({range_label})"] = df[f"{calculation_mode} ({range_label})"] * 2
#     df[f"Hologram ({range_label})"] = df["Total Candidate"].apply(lambda x: math.ceil(x / 100) + 1)
#     df[f"Id Card ({range_label})"] = df[f"All-Day Max Opr ({range_label})"] + 1
#     df[f"Jacket ({range_label})"] = df[f"All-Day Max Opr ({range_label})"]
#     df["Range"] = range_label
#     return df
#
#
# # Streamlit App
# st.title("Operational Ratio Processing")
#
# # File upload
# uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
#
# if uploaded_file:
#     # Load the file into a DataFrame
#     df = pd.read_excel(uploaded_file)
#     df.fillna(0, inplace=True)
#
#     # Allow the user to select columns for day headers
#     st.header("Select Columns for Day Headers")
#     day_headers = {}
#     for day in ["Day-1"]:  # You can add more days if needed
#         selected_cols = st.multiselect(f"Select columns for {day}", df.columns.tolist())
#         if selected_cols:
#             day_headers[day] = selected_cols
#
#     # Ensure day_headers is not empty
#     if not day_headers:
#         st.warning("Please select at least one column for day headers.")
#         st.stop()
#
#     # Allow the user to choose between IRIS and FPS
#     calculation_mode = st.radio("Choose Calculation Mode (Mandatory)", options=["IRIS", "FPS"])
#
#     # Validate the selection of the calculation mode
#     if not calculation_mode:
#         st.warning("You must select either IRIS or FPS to proceed.")
#         st.stop()
#
#     # Default and customizable ranges
#     st.header("Select Ranges")
#     default_ranges = [(61, 61), (66, 66), (71, 71), (76, 76), (81, 81), (86, 86), (91, 91), (96, 96), (101, 101)]
#     default_ranges_display = ", ".join([f"({min_val}, {max_val})" for min_val, max_val in default_ranges])
#     st.write(f"Default Ranges: {default_ranges_display}")
#
#     custom_ranges_input = st.text_input(
#         "Enter custom ranges (e.g., 50-55, 60-65). Leave blank to use default ranges."
#     )
#     if custom_ranges_input.strip():
#         try:
#             custom_ranges = [
#                 tuple(map(int, r.strip().split("-"))) for r in custom_ranges_input.split(",")
#             ]
#         except ValueError:
#             st.error("Invalid range format. Please use the format 'min-max' separated by commas.")
#             st.stop()
#     else:
#         custom_ranges = default_ranges
#
#     # Calculate "Total Candidate"
#     df["Total Candidate"] = df[[header for headers in day_headers.values() for header in headers]].sum(axis=1)
#
#     # Process ranges
#     detailed_results = []
#     grand_totals_combined = []
#
#     for min_val, max_val in custom_ranges:
#         label = f"{min_val}-{max_val}"
#         processed_df = process_range(df.copy(), min_val, max_val, label, day_headers, calculation_mode)
#         detailed_results.append((label, processed_df))
#
#         totals = {
#             "Range": label,
#             **{f"{header} Opr": processed_df[f"{header} Opr ({label})"].sum() for day, headers in day_headers.items()
#                for header in headers},
#             "All-Day Max Opr": processed_df[f"All-Day Max Opr ({label})"].sum(),
#             "Tab": processed_df[f"Tab ({label})"].sum(),
#             "Master Tab": processed_df[f"Master Tab ({label})"].sum(),
#             "SIM": processed_df[f"SIM ({label})"].sum(),
#             f"{calculation_mode}": processed_df[f"{calculation_mode} ({label})"].sum(),
#             "OTG": processed_df[f"OTG ({label})"].sum(),
#             "Hologram": processed_df[f"Hologram ({label})"].sum(),
#             "Id Card": processed_df[f"Id Card ({label})"].sum(),
#             "Jacket": processed_df[f"Jacket ({label})"].sum(),
#         }
#         grand_totals_combined.append(totals)
#
#     # Create Grand Totals DataFrame
#     grand_totals_final = pd.DataFrame(grand_totals_combined)
#
#     # Save to Excel in-memory
#     output = io.BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
#         for label, detailed_df in detailed_results:
#             if calculation_mode == "FPS":
#                 detailed_df = detailed_df.drop(columns=[col for col in detailed_df.columns if "IRIS" in col])
#             detailed_df.to_excel(writer, sheet_name=label, index=False)
#         grand_totals_final.to_excel(writer, sheet_name="Grand Totals", index=False)
#
#     # Button to download the result
#     st.download_button(
#         label="Download Processed Excel File",
#         data=output.getvalue(),
#         file_name="processed_data.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     )















import streamlit as st
import pandas as pd
import math
import io

# Functions for calculations
def calculate_opr(value, min_value, max_value):
    if value == 0:
        return 0
    if value <= min_value:
        return 1
    opr = math.ceil(value / min_value)
    if opr > 1 and value / (opr - 1) < max_value:
        return opr - 1
    return opr


def calculate_tab(max_opr, master_tab=1):
    if max_opr == 0:
        return 0
    if 8 <= max_opr < 16:
        return max_opr + 2 - master_tab
    elif max_opr > 15:
        return max_opr + 3 - master_tab
    else:
        return max_opr + 1 - master_tab


def process_range(df, min_value, max_value, range_label, day_headers, calculation_mode):
    for day, headers in day_headers.items():
        for header in headers:
            opr_col = f"{header} Opr ({range_label})"
            df[opr_col] = df[header].apply(lambda x: calculate_opr(x, min_value, max_value))

    opr_columns = [f"{header} Opr ({range_label})" for day, headers in day_headers.items() for header in headers]
    df[f"All-Day Max Opr ({range_label})"] = df[opr_columns].max(axis=1)

    df[f"Tab ({range_label})"] = df[f"All-Day Max Opr ({range_label})"].apply(lambda x: calculate_tab(x))
    df[f"Master Tab ({range_label})"] = 1
    df[f"SIM ({range_label})"] = 1

    if calculation_mode == "IRIS":
        df[f"IRIS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]
        df[f"FPS ({range_label})"] = 1
    elif calculation_mode == "FPS":
        df[f"FPS ({range_label})"] = df[f"Tab ({range_label})"] + df[f"Master Tab ({range_label})"]

    df[f"OTG ({range_label})"] = df[f"{calculation_mode} ({range_label})"] * 2
    df[f"Hologram ({range_label})"] = df["Total Candidate"].apply(lambda x: math.ceil(x / 100) + 1)
    df[f"Id Card ({range_label})"] = df[f"All-Day Max Opr ({range_label})"] + 1
    df[f"Jacket ({range_label})"] = df[f"All-Day Max Opr ({range_label})"]
    df["Range"] = range_label
    return df


# Streamlit App
st.set_page_config(
    page_title="Operational Ratio Processing",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Sidebar
with st.sidebar:
    st.title("Operational Ratio Processing 📊")
    st.markdown("""
    This app processes operational ratios for IRIS or FPS services.
    Follow these steps:
    1. **Upload your Excel file.**
    2. **Select day headers for processing.**
    3. **Choose a calculation mode (IRIS or FPS).**
    4. **Customize ranges or use defaults.**
    5. **Download your processed file.**
    """)
    st.markdown("---")
    st.write("📌 **Made by Vikrant Kumar (DM)**")
    st.markdown("""
        For feedback or queries:
        - Email: vikrant.kumar@example.com
        - GitHub: [vikrant-dm](https://github.com/vikrant-dm)
    """)

# Main Interface
st.title("Operational Ratio Processor 📈")
st.markdown("""
Welcome to the **Operational Ratio Processing App**. 
This tool allows you to process operational data and generate detailed insights.
""")
st.markdown("---")

# File upload
uploaded_file = st.file_uploader("Upload your Excel file 📄", type=["xlsx"], help="Upload an Excel file containing your operational data.")

if uploaded_file:
    # Load the file into a DataFrame
    df = pd.read_excel(uploaded_file)
    df.fillna(0, inplace=True)

    # Select columns for day headers
    st.header("Step 1: Select Columns for Day Headers")
    day_headers = {}
    for day in ["Day-1"]:  # Add more days if needed
        selected_cols = st.multiselect(f"Select columns for {day} 📅", df.columns.tolist())
        if selected_cols:
            day_headers[day] = selected_cols

    # Ensure day_headers is not empty
    if not day_headers:
        st.warning("⚠️ Please select at least one column for day headers to proceed.")
        st.stop()

    # Choose calculation mode
    st.header("Step 2: Choose Calculation Mode")
    calculation_mode = st.radio("Select a service (mandatory):", options=["IRIS", "FPS"], help="Select the service for which you want to process the data.")

    if not calculation_mode:
        st.warning("⚠️ You must select either IRIS or FPS to proceed.")
        st.stop()

    # Range selection
    st.header("Step 3: Select Ranges")
    default_ranges = [(61, 61), (66, 66), (71, 71), (76, 76), (81, 81), (86, 86), (91, 91), (96, 96), (101, 101)]
    default_ranges_display = ", ".join([f"({min_val}, {max_val})" for min_val, max_val in default_ranges])
    st.write(f"Default Ranges: {default_ranges_display}")

    custom_ranges_input = st.text_input(
        "Enter custom ranges (e.g., 50-55, 60-65) or leave blank for defaults."
    )
    if custom_ranges_input.strip():
        try:
            custom_ranges = [
                tuple(map(int, r.strip().split("-"))) for r in custom_ranges_input.split(",")
            ]
        except ValueError:
            st.error("Invalid range format. Please use the format 'min-max' separated by commas.")
            st.stop()
    else:
        custom_ranges = default_ranges

    # Calculate "Total Candidate"
    df["Total Candidate"] = df[[header for headers in day_headers.values() for header in headers]].sum(axis=1)

    # Process ranges
    st.header("Step 4: Processing Data")
    detailed_results = []
    grand_totals_combined = []

    for min_val, max_val in custom_ranges:
        label = f"{min_val}-{max_val}"
        processed_df = process_range(df.copy(), min_val, max_val, label, day_headers, calculation_mode)
        detailed_results.append((label, processed_df))

        totals = {
            "Range": label,
            **{f"{header} Opr": processed_df[f"{header} Opr ({label})"].sum() for day, headers in day_headers.items()
               for header in headers},
            "All-Day Max Opr": processed_df[f"All-Day Max Opr ({label})"].sum(),
            "Tab": processed_df[f"Tab ({label})"].sum(),
            "Master Tab": processed_df[f"Master Tab ({label})"].sum(),
            "SIM": processed_df[f"SIM ({label})"].sum(),
            f"{calculation_mode}": processed_df[f"{calculation_mode} ({label})"].sum(),
            "OTG": processed_df[f"OTG ({label})"].sum(),
            "Hologram": processed_df[f"Hologram ({label})"].sum(),
            "Id Card": processed_df[f"Id Card ({label})"].sum(),
            "Jacket": processed_df[f"Jacket ({label})"].sum(),
        }
        grand_totals_combined.append(totals)

    # Create Grand Totals DataFrame
    grand_totals_final = pd.DataFrame(grand_totals_combined)

    # Save to Excel in-memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for label, detailed_df in detailed_results:
            if calculation_mode == "FPS":
                detailed_df = detailed_df.drop(columns=[col for col in detailed_df.columns if "IRIS" in col])
            detailed_df.to_excel(writer, sheet_name=label, index=False)
        grand_totals_final.to_excel(writer, sheet_name="Grand Totals", index=False)

    # Button to download the result
    st.success("✅ Processing Complete!")
    st.download_button(
        label="📥 Download Processed Excel File",
        data=output.getvalue(),
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Footer
st.markdown("---")
st.markdown("📌 **Made by Vikrant Kumar (DM)**")

