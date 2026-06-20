import streamlit as st
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import time


@st.cache_data
def parse_datumaro_json(json_data):
    """
    Parse Datumaro JSON format and extract annotation data
    Returns: list of dicts with image_name, frame_number, and label annotations
    """
    try:
        data = json.loads(json_data) if isinstance(json_data, str) else json_data
        
        # Build label_id to label_name mapping from categories
        label_map = {}
        if "categories" in data and "label" in data["categories"]:
            labels = data["categories"]["label"].get("labels", [])
            for idx, label_obj in enumerate(labels):
                label_map[idx] = label_obj.get("name", f"Label_{idx}")
        
        # Extract items from Datumaro format
        items = data.get("items", [])
        annotations_data = []
        
        for frame_number, item in enumerate(items, start=0):
            image_name = item.get("id", "")
            
            # Extract annotations for this image
            annotations = item.get("annotations", [])
            label_counts = {}
            
            for annotation in annotations:
                label_id = annotation.get("label_id", -1)
                label_name = label_map.get(label_id, f"Unknown_{label_id}")
                annotation_type = annotation.get("type", "")
                
                if label_name not in label_counts:
                    label_counts[label_name] = {"box": 0, "polygon": 0}
                
                # Count by type
                if annotation_type == "bbox":
                    label_counts[label_name]["box"] += 1
                elif annotation_type == "polygon":
                    label_counts[label_name]["polygon"] += 1
            
            annotations_data.append({
                "image_name": image_name,
                "frame_number": frame_number,
                "label_counts": label_counts
            })
        
        return annotations_data
    
    except json.JSONDecodeError as e:
        st.error(f"Invalid JSON format: {str(e)}")
        return None
    except Exception as e:
        st.error(f"Error parsing JSON: {str(e)}")
        return None


@st.cache_data
def create_excel_with_labels(json_data):
    """
    Create Excel file with dynamic label columns
    Structure: SI No | Image Name | Frame Number | [Label 1 Box | Label 1 Polygon] | ... | Total Annotations
    """
    annotations_data = parse_datumaro_json(json_data)
    if not annotations_data:
        return None, None
    
    # Collect all unique labels
    all_labels = set()
    for item in annotations_data:
        all_labels.update(item["label_counts"].keys())
    all_labels = sorted(list(all_labels))
    
    # Prepare data for DataFrame
    rows = []
    for si, item in enumerate(annotations_data, 1):
        row = {
            "SI No": si,
            "Image Name": item["image_name"],
            "Frame Number": item["frame_number"]
        }
        
        total_annotations = 0
        for label in all_labels:
            box_count = item["label_counts"].get(label, {}).get("box", 0)
            polygon_count = item["label_counts"].get(label, {}).get("polygon", 0)
            
            row[f"{label} - Box"] = box_count
            row[f"{label} - Polygon"] = polygon_count
            total_annotations += box_count + polygon_count
        
        row["Total Annotations"] = total_annotations
        rows.append(row)
    
    df = pd.DataFrame(rows)
    
    # Create Excel workbook with formatting
    wb = Workbook()
    ws = wb.active
    ws.title = "Annotations"
    
    # Write headers
    headers = ["SI No", "Image Name", "Frame Number"]
    for label in all_labels:
        headers.append(f"{label} - Box")
        headers.append(f"{label} - Polygon")
    headers.append("Total Annotations")
    
    ws.append(headers)
    
    # Format header row
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    # Add data rows
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row_data in rows:
        row_values = [
            row_data["SI No"],
            row_data["Image Name"],
            row_data["Frame Number"]
        ]
        for label in all_labels:
            row_values.append(row_data[f"{label} - Box"])
            row_values.append(row_data[f"{label} - Polygon"])
        row_values.append(row_data["Total Annotations"])
        
        ws.append(row_values)
    
    # Apply borders and alignment to all cells
    for row in ws.iter_rows(min_row=2, max_row=len(rows)+1):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15
    for col in range(4, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14
    
    # Add summary row with totals
    summary_row = len(rows) + 3
    ws[f'A{summary_row}'] = "SUMMARY - Total Counts:"
    ws[f'A{summary_row}'].font = Font(bold=True)
    
    # Add summary formulas for ALL columns (from D onwards to last column)
    for col_idx in range(4, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        cell = ws[f'{col_letter}{summary_row}']
        cell.value = f'=SUM({col_letter}2:{col_letter}{len(rows)+1})'
        cell.font = Font(bold=True)
        
        # Color coding: Light green for label columns, light red for total
        if col_idx == len(headers):  # Total Annotations column
            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        else:
            cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    return wb, df


@st.cache_data
def parse_polygon_points_data(json_data):
    """
    Parse Datumaro JSON and extract per-polygon point data.
    Returns list of dicts: image_name, label_name, polygon_id, points_count
    """
    try:
        data = json.loads(json_data) if isinstance(json_data, str) else json_data

        label_map = {}
        if "categories" in data and "label" in data["categories"]:
            labels = data["categories"]["label"].get("labels", [])
            for idx, label_obj in enumerate(labels):
                label_map[idx] = label_obj.get("name", f"Label_{idx}")

        items = data.get("items", [])
        polygon_rows = []

        for item in items:
            image_name = item.get("id", "")
            annotations = item.get("annotations", [])

            for annotation in annotations:
                if annotation.get("type", "") == "polygon":
                    label_id = annotation.get("label_id", -1)
                    label_name = label_map.get(label_id, f"Unknown_{label_id}")
                    polygon_id = annotation.get("id", "N/A")
                    points = annotation.get("points", [])
                    points_count = len(points) // 2

                    polygon_rows.append({
                        "image_name": image_name,
                        "label_name": label_name,
                        "polygon_id": polygon_id,
                        "points_count": points_count
                    })

        return polygon_rows

    except json.JSONDecodeError as e:
        st.error(f"Invalid JSON format: {str(e)}")
        return None
    except Exception as e:
        st.error(f"Error parsing JSON: {str(e)}")
        return None


def annotation_converter_page():
    st.title("🔄 Datumaro JSON to Excel Converter")

    uploaded_file = st.file_uploader(
        "Choose a Datumaro JSON file",
        type="json",
        key="converter_uploader"
    )

    if uploaded_file is not None:
        try:
            json_content = uploaded_file.read().decode("utf-8")

            msg_placeholder = st.empty()
            msg_placeholder.info("📋 Processing JSON file...")

            wb, df = create_excel_with_labels(json_content)

            if df is None or df.empty:
                st.error("No valid annotation data found in JSON file")
                return

            msg_placeholder.success(f"✅ Found {len(df)} images with annotations")
            time.sleep(3)
            msg_placeholder.empty()

            df['Job'] = df['Image Name'].str.split('_').str[0]

            label_cols = [col for col in df.columns if "- Box" in col]
            selected_labels = sorted([col.replace(" - Box", "") for col in label_cols])

            filtered_df = df.copy()

            st.divider()

            col_stats, col_download = st.columns([2, 1])

            with col_stats:
                st.subheader("📈 Statistics")
                stat_col1, stat_col2, stat_col3 = st.columns(3)

                with stat_col1:
                    st.metric("Total Images", len(filtered_df))

                with stat_col2:
                    st.metric("Total Annotations", int(filtered_df["Total Annotations"].sum()))

                with stat_col3:
                    st.metric("Unique Labels", len(selected_labels))

            with col_download:
                st.subheader("💾 Download")

                output_buffer = io.BytesIO()

                filtered_df_for_export = filtered_df.copy()
                wb_filtered = Workbook()
                ws = wb_filtered.active
                ws.title = "Annotations"

                headers = ["SI No", "Job", "Image Name", "Frame Number"]
                for label in selected_labels:
                    headers.append(f"{label} - Box")
                    headers.append(f"{label} - Polygon")
                headers.append("Total Annotations of the Image")

                ws.append(headers)

                header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment

                border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )

                for idx, row_data in filtered_df_for_export.iterrows():
                    row_values = [
                        row_data["SI No"],
                        row_data["Job"],
                        row_data["Image Name"],
                        row_data["Frame Number"]
                    ]
                    for label in selected_labels:
                        row_values.append(row_data[f"{label} - Box"])
                        row_values.append(row_data[f"{label} - Polygon"])
                    row_values.append(row_data["Total Annotations"])

                    ws.append(row_values)

                for row in ws.iter_rows(min_row=2, max_row=len(filtered_df_for_export)+1):
                    for cell in row:
                        cell.border = border
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                ws.column_dimensions['A'].width = 8
                ws.column_dimensions['B'].width = 15
                ws.column_dimensions['C'].width = 25
                ws.column_dimensions['D'].width = 15
                for col in range(5, len(headers) + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 14

                summary_row = len(filtered_df_for_export) + 3
                ws[f'A{summary_row}'] = "SUMMARY - Total Counts:"
                ws[f'A{summary_row}'].font = Font(bold=True)

                for col_idx in range(5, len(headers) + 1):
                    col_letter = get_column_letter(col_idx)
                    cell = ws[f'{col_letter}{summary_row}']
                    cell.value = f'=SUM({col_letter}2:{col_letter}{len(filtered_df_for_export)+1})'
                    cell.font = Font(bold=True)

                    if col_idx == len(headers):
                        cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    else:
                        cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

                wb_filtered.save(output_buffer)
                output_buffer.seek(0)

                filename = f"{uploaded_file.name.split('.')[0]}_annotations_filtered.xlsx"

                st.download_button(
                    label="📥 Download Excel",
                    data=output_buffer.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.divider()

            st.subheader("📊 Data Preview")
            col_order = ["SI No", "Job", "Image Name", "Frame Number"]
            for label in selected_labels:
                col_order.append(f"{label} - Box")
                col_order.append(f"{label} - Polygon")
            col_order.append("Total Annotations")

            display_df = filtered_df[col_order]
            st.dataframe(display_df, use_container_width=True, height=500, hide_index=True)

            st.divider()
            st.subheader("🏷️ Annotation Breakdown by Label")

            if selected_labels:
                label_stats = {}
                for label in selected_labels:
                    box_col = f"{label} - Box"
                    polygon_col = f"{label} - Polygon"

                    box_count = filtered_df[box_col].sum()
                    polygon_count = filtered_df[polygon_col].sum()

                    label_stats[label] = {
                        "Box": int(box_count),
                        "Polygon": int(polygon_count),
                        "Total": int(box_count + polygon_count)
                    }

                stats_df = pd.DataFrame(label_stats).T
                st.dataframe(stats_df, use_container_width=True)

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            import traceback
            st.error(traceback.format_exc())


def polygon_point_counter_page():
    st.title("🔷 Polygon Point Counter")
    st.markdown("Analyze per-polygon vertex counts from Datumaro JSON files.")

    uploaded_file = st.file_uploader(
        "Choose a Datumaro JSON file",
        type="json",
        key="polygon_uploader"
    )

    if uploaded_file is not None:
        try:
            json_content = uploaded_file.read().decode("utf-8")

            msg_placeholder = st.empty()
            msg_placeholder.info("🔷 Extracting polygon point data...")

            polygon_rows = parse_polygon_points_data(json_content)

            if not polygon_rows:
                st.error("No polygon annotations found in the JSON file.")
                return

            msg_placeholder.success(f"✅ Found {len(polygon_rows)} polygons")
            time.sleep(2)
            msg_placeholder.empty()

            # Extract unique labels for checkbox selection
            all_labels = sorted(list(set(row["label_name"] for row in polygon_rows)))

            # Add checkbox section for label selection
            st.subheader("🏷️ Select Labels to Include")
            selected_labels = st.multiselect(
                "Choose which labels to include in the analysis and download:",
                options=all_labels,
                default=all_labels,
                key="label_filter"
            )

            # Validation: show warning if no labels selected
            if not selected_labels:
                st.warning("⚠️ Please select at least one label to proceed.")
                return

            # Filter polygon_rows based on selected labels
            filtered_polygon_rows = [row for row in polygon_rows if row["label_name"] in selected_labels]

            df = pd.DataFrame(filtered_polygon_rows)[["image_name", "label_name", "points_count"]].copy()
            df.insert(0, "SI No", range(1, len(df) + 1))
            df.columns = ["SI No", "Image Name", "Label Name", "Points Count"]

            st.divider()

            col_stats, col_download = st.columns([2, 1])

            with col_stats:
                st.subheader("📈 Statistics")
                s1, s2 = st.columns(2)
                with s1:
                    st.metric("Total Polygons", len(df))
                with s2:
                    st.metric("Images with Polygons", df["Image Name"].nunique())

            with col_download:
                st.subheader("💾 Download")

                wb_poly = Workbook()
                ws = wb_poly.active
                ws.title = "Polygon Points"

                poly_headers = ["SI No", "Image Name", "Label Name", "Points Count"]
                ws.append(poly_headers)

                header_fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment

                border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )

                for _, row_data in df.iterrows():
                    ws.append([
                        int(row_data["SI No"]),
                        row_data["Image Name"],
                        row_data["Label Name"],
                        int(row_data["Points Count"])
                    ])

                for row in ws.iter_rows(min_row=2, max_row=len(df) + 1):
                    for cell in row:
                        cell.border = border
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                for row in ws.iter_rows(min_row=2, max_row=len(df) + 1, min_col=2, max_col=2):
                    for cell in row:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                ws.column_dimensions['A'].width = 8
                ws.column_dimensions['B'].width = 30
                ws.column_dimensions['C'].width = 20
                ws.column_dimensions['D'].width = 15

                summary_row = len(df) + 3
                ws[f'A{summary_row}'] = "TOTAL POLYGONS:"
                ws[f'A{summary_row}'].font = Font(bold=True)
                poly_total_cell = ws[f'C{summary_row}']
                poly_total_cell.value = len(df)
                poly_total_cell.font = Font(bold=True)
                poly_total_cell.fill = PatternFill(start_color="E2D0F0", end_color="E2D0F0", fill_type="solid")

                ws[f'A{summary_row + 1}'] = "TOTAL POINTS:"
                ws[f'A{summary_row + 1}'].font = Font(bold=True)
                pts_total_cell = ws[f'D{summary_row + 1}']
                pts_total_cell.value = f'=SUM(D2:D{len(df) + 1})'
                pts_total_cell.font = Font(bold=True)
                pts_total_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

                ws2 = wb_poly.create_sheet("Label Summary")
                sum_headers = ["Label Name", "Polygon Count", "Total Points"]
                ws2.append(sum_headers)

                for cell in ws2[1]:
                    cell.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                label_groups = {}
                for row in filtered_polygon_rows:
                    ln = row["label_name"]
                    if ln not in label_groups:
                        label_groups[ln] = {"count": 0, "total_points": 0}
                    label_groups[ln]["count"] += 1
                    label_groups[ln]["total_points"] += row["points_count"]

                for label, stats in sorted(label_groups.items()):
                    ws2.append([label, stats["count"], stats["total_points"]])

                for row in ws2.iter_rows(min_row=2, max_row=len(label_groups) + 1):
                    for cell in row:
                        cell.border = border
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                ws2.column_dimensions['A'].width = 20
                ws2.column_dimensions['B'].width = 15
                ws2.column_dimensions['C'].width = 15

                output_buffer = io.BytesIO()
                wb_poly.save(output_buffer)
                output_buffer.seek(0)

                filename = f"{uploaded_file.name.split('.')[0]}_polygon_points.xlsx"

                st.download_button(
                    label="📥 Download Excel",
                    data=output_buffer.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.divider()

            st.subheader("🔷 Polygon Points Data")
            st.dataframe(df, use_container_width=True, height=500, hide_index=True)

            st.divider()
            st.subheader("📊 Per-Label Summary")

            label_summary_df = df.groupby("Label Name").agg(
                Polygon_Count=("Points Count", "count"),
                Total_Points=("Points Count", "sum")
            ).reset_index()
            label_summary_df.columns = ["Label Name", "Polygon Count", "Total Points"]
            st.dataframe(label_summary_df, use_container_width=True, hide_index=True)

            st.divider()
            total_col1, total_col2 = st.columns(2)
            with total_col1:
                st.metric("🔢 Total Polygons", len(df))
            with total_col2:
                st.metric("📍 Total Points (All Polygons)", int(df["Points Count"].sum()))

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            import traceback
            st.error(traceback.format_exc())


def filter_json_by_labels(json_data, selected_labels, keep_selected=True):
    """
    Filter Datumaro JSON to keep only annotations for selected or non-selected labels.
    Args:
        json_data: Datumaro JSON as string or dict
        selected_labels: List of label names to filter by
        keep_selected: If True, keep selected labels; if False, keep non-selected labels
    Returns filtered JSON as string.
    """
    try:
        data = json.loads(json_data) if isinstance(json_data, str) else json_data

        # Get label_id to label_name mapping
        label_map = {}
        if "categories" in data and "label" in data["categories"]:
            labels = data["categories"]["label"].get("labels", [])
            for idx, label_obj in enumerate(labels):
                label_map[idx] = label_obj.get("name", f"Label_{idx}")

        # Create reverse mapping: label_name to label_id
        name_to_id = {v: k for k, v in label_map.items()}

        # Filter items: keep only annotations for selected or non-selected labels
        filtered_items = []
        for item in data.get("items", []):
            filtered_annotations = []
            for annotation in item.get("annotations", []):
                label_id = annotation.get("label_id", -1)
                label_name = label_map.get(label_id, f"Unknown_{label_id}")
                if keep_selected:
                    # Keep selected labels
                    if label_name in selected_labels:
                        filtered_annotations.append(annotation)
                else:
                    # Keep non-selected labels
                    if label_name not in selected_labels:
                        filtered_annotations.append(annotation)

            # Only keep item if it has annotations after filtering
            if filtered_annotations:
                item_copy = item.copy()
                item_copy["annotations"] = filtered_annotations
                filtered_items.append(item_copy)

        # Update categories to include only selected or non-selected labels
        if "categories" in data and "label" in data["categories"]:
            filtered_labels = []
            new_label_map = {}
            new_id = 0
            for label_obj in data["categories"]["label"].get("labels", []):
                label_name = label_obj.get("name", "")
                if keep_selected:
                    # Keep selected labels
                    if label_name in selected_labels:
                        new_label_map[new_id] = label_name
                        label_obj_copy = label_obj.copy()
                        label_obj_copy["id"] = new_id
                        filtered_labels.append(label_obj_copy)
                        new_id += 1
                else:
                    # Keep non-selected labels
                    if label_name not in selected_labels:
                        new_label_map[new_id] = label_name
                        label_obj_copy = label_obj.copy()
                        label_obj_copy["id"] = new_id
                        filtered_labels.append(label_obj_copy)
                        new_id += 1

            data["categories"]["label"]["labels"] = filtered_labels

            # Update label_ids in annotations to match new mapping
            old_to_new_id = {}
            for old_id, name in label_map.items():
                if keep_selected:
                    if name in selected_labels:
                        for new_id, new_name in new_label_map.items():
                            if new_name == name:
                                old_to_new_id[old_id] = new_id
                                break
                else:
                    if name not in selected_labels:
                        for new_id, new_name in new_label_map.items():
                            if new_name == name:
                                old_to_new_id[old_id] = new_id
                                break

            for item in filtered_items:
                for annotation in item["annotations"]:
                    old_label_id = annotation.get("label_id", -1)
                    if old_label_id in old_to_new_id:
                        annotation["label_id"] = old_to_new_id[old_label_id]

        data["items"] = filtered_items
        return json.dumps(data, indent=2)

    except Exception as e:
        raise Exception(f"Error filtering JSON: {str(e)}")


def merge_datumaro_jsons(json_files_list):
    """
    Merge multiple Datumaro JSON files into one.
    Returns merged JSON as string.
    Raises Exception if duplicate label names are found across files.
    """
    try:
        all_data = []
        for json_content in json_files_list:
            data = json.loads(json_content) if isinstance(json_content, str) else json_content
            all_data.append(data)

        # Check for duplicate label names across all files
        all_labels_with_file = []
        for idx, data in enumerate(all_data):
            if "categories" in data and "label" in data["categories"]:
                for label_obj in data["categories"]["label"].get("labels", []):
                    label_name = label_obj.get("name", "")
                    all_labels_with_file.append((label_name, idx))

        # Find duplicates
        label_names = [label[0] for label in all_labels_with_file]
        unique_labels = set(label_names)
        if len(label_names) != len(unique_labels):
            # Find which labels are duplicated
            from collections import Counter
            label_counts = Counter(label_names)
            duplicates = {label: count for label, count in label_counts.items() if count > 1}

            # Build error message with file information
            error_msg = "Duplicate label names found across files:\n"
            for dup_label in duplicates:
                files_with_label = [f"File {idx+1}" for label, idx in all_labels_with_file if label == dup_label]
                error_msg += f"- '{dup_label}' appears in: {', '.join(files_with_label)}\n"
            error_msg += "\nPlease remove or rename duplicate labels before merging."
            raise Exception(error_msg)

        # Merge categories: combine all unique labels
        merged_labels = []
        label_name_to_new_id = {}
        new_id = 0

        for data in all_data:
            if "categories" in data and "label" in data["categories"]:
                for label_obj in data["categories"]["label"].get("labels", []):
                    label_name = label_obj.get("name", f"Label_{new_id}")
                    if label_name not in label_name_to_new_id:
                        label_name_to_new_id[label_name] = new_id
                        merged_labels.append({
                            "id": new_id,
                            "name": label_name,
                            **{k: v for k, v in label_obj.items() if k not in ["id", "name"]}
                        })
                        new_id += 1

        # Build old label_id to new label_id mapping for each file
        file_mappings = []
        for data in all_data:
            old_to_new = {}
            if "categories" in data and "label" in data["categories"]:
                for label_obj in data["categories"]["label"].get("labels", []):
                    old_id = label_obj.get("id", -1)
                    label_name = label_obj.get("name", "")
                    if label_name in label_name_to_new_id:
                        old_to_new[old_id] = label_name_to_new_id[label_name]
            file_mappings.append(old_to_new)

        # Merge items: combine all items, merging annotations for same image_id
        merged_items = {}
        for idx, data in enumerate(all_data):
            mapping = file_mappings[idx]
            for item in data.get("items", []):
                image_id = item.get("id", "")
                
                if image_id not in merged_items:
                    merged_items[image_id] = {
                        "id": image_id,
                        "annotations": []
                    }
                
                # Add annotations with updated label_ids
                for annotation in item.get("annotations", []):
                    annotation_copy = annotation.copy()
                    old_label_id = annotation_copy.get("label_id", -1)
                    if old_label_id in mapping:
                        annotation_copy["label_id"] = mapping[old_label_id]
                    merged_items[image_id]["annotations"].append(annotation_copy)

        # Build merged JSON structure
        merged_data = {
            "categories": {
                "label": {
                    "labels": merged_labels
                }
            },
            "items": list(merged_items.values())
        }

        return json.dumps(merged_data, indent=2)

    except Exception as e:
        raise Exception(f"Error merging JSONs: {str(e)}")


def json_split_merge_page():
    st.title("🔀 JSON Split & Merge")
    st.markdown("Filter annotations by label or merge multiple JSON files.")

    # Mode selection
    mode = st.radio(
        "Select Mode:",
        ["🔍 Split/Filter by Labels", "🔗 Merge Multiple JSONs"],
        label_visibility="collapsed",
        horizontal=True
    )

    st.divider()

    if mode == "🔍 Split/Filter by Labels":
        st.subheader("🔍 Filter JSON by Labels")
        st.markdown("Upload a Datumaro JSON file and select which labels to keep.")

        uploaded_file = st.file_uploader(
            "Choose a Datumaro JSON file",
            type="json",
            key="filter_uploader"
        )

        if uploaded_file is not None:
            try:
                json_content = uploaded_file.read().decode("utf-8")
                data = json.loads(json_content)

                # Extract all unique labels
                all_labels = []
                if "categories" in data and "label" in data["categories"]:
                    for label_obj in data["categories"]["label"].get("labels", []):
                        all_labels.append(label_obj.get("name", f"Unknown"))
                all_labels = sorted(list(set(all_labels)))

                if not all_labels:
                    st.error("No labels found in the JSON file.")
                    return

                st.info(f"Found {len(all_labels)} labels in the file.")

                # Label selection
                st.subheader("🏷️ Select Labels to Keep")
                selected_labels = st.multiselect(
                    "Choose which labels to include in the filtered JSON:",
                    options=all_labels,
                    default=all_labels,
                    key="filter_labels"
                )

                if not selected_labels:
                    st.warning("⚠️ Please select at least one label to proceed.")
                    return

                # Filter and preview
                if st.button("🔄 Filter JSON", use_container_width=True):
                    with st.spinner("Filtering JSON..."):
                        # Filter selected labels
                        filtered_json_selected = filter_json_by_labels(json_content, selected_labels, keep_selected=True)
                        filtered_data_selected = json.loads(filtered_json_selected)

                        # Filter non-selected labels
                        other_labels = [label for label in all_labels if label not in selected_labels]
                        filtered_json_other = filter_json_by_labels(json_content, selected_labels, keep_selected=False)
                        filtered_data_other = json.loads(filtered_json_other)

                        st.success(f"✅ Filtered JSONs created")

                        # Show statistics side by side with vertical divider
                        col_stats_left, col_divider, col_stats_right = st.columns([1, 0.05, 1])

                        with col_stats_left:
                            st.subheader("📊 Filtered JSON Statistics")
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Total Items", len(filtered_data_selected["items"]))
                            with col2:
                                total_annotations = sum(len(item.get("annotations", [])) for item in filtered_data_selected["items"])
                                st.metric("Total Annotations", total_annotations)
                            with col3:
                                st.metric("Labels Kept", len(selected_labels))

                        with col_divider:
                            st.markdown("""
                                <style>
                                    .vertical-line {
                                        border-left: 2px solid #e0e0e0;
                                        height: 100%;
                                        margin: 0 auto;
                                    }
                                </style>
                                <div class="vertical-line"></div>
                            """, unsafe_allow_html=True)

                        if other_labels:
                            with col_stats_right:
                                st.subheader("📊 Default JSON Statistics")
                                col4, col5, col6 = st.columns(3)
                                with col4:
                                    st.metric("Total Items", len(filtered_data_other["items"]))
                                with col5:
                                    total_annotations_other = sum(len(item.get("annotations", [])) for item in filtered_data_other["items"])
                                    st.metric("Total Annotations", total_annotations_other)
                                with col6:
                                    st.metric("Labels Kept", len(other_labels))

                        st.divider()

                        # Download buttons
                        st.subheader("💾 Download Filtered JSONs")
                        col_dl1, col_dl2 = st.columns(2)

                        with col_dl1:
                            filename_filtered = f"{uploaded_file.name.split('.')[0]}_filtered.json"
                            st.download_button(
                                label="📥 Download Filtered JSON",
                                data=filtered_json_selected,
                                file_name=filename_filtered,
                                mime="application/json",
                                use_container_width=True
                            )

                        if other_labels:
                            with col_dl2:
                                filename_default = f"{uploaded_file.name.split('.')[0]}_default.json"
                                st.download_button(
                                    label="📥 Download Default JSON",
                                    data=filtered_json_other,
                                    file_name=filename_default,
                                    mime="application/json",
                                    use_container_width=True
                                )

            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
                import traceback
                st.error(traceback.format_exc())

    else:  # Merge mode
        st.subheader("🔗 Merge Multiple JSON Files")
        st.markdown("Upload multiple Datumaro JSON files to merge them into one.")

        uploaded_files = st.file_uploader(
            "Choose Datumaro JSON files (2 or more)",
            type="json",
            accept_multiple_files=True,
            key="merge_uploader"
        )

        if uploaded_files is not None:
            if len(uploaded_files) < 2:
                st.warning("⚠️ Please upload at least 2 JSON files to merge.")
                return

            st.info(f"📁 {len(uploaded_files)} files selected for merging.")

            # Show file summary
            st.subheader("📋 Files to Merge")
            for idx, file in enumerate(uploaded_files, 1):
                st.write(f"{idx}. {file.name}")

            if st.button("🔗 Merge JSONs", use_container_width=True):
                try:
                    with st.spinner("Merging JSON files..."):
                        json_contents = []
                        for file in uploaded_files:
                            content = file.read().decode("utf-8")
                            json_contents.append(content)

                        merged_json = merge_datumaro_jsons(json_contents)
                        merged_data = json.loads(merged_json)

                        st.success(f"✅ Successfully merged {len(uploaded_files)} JSON files")

                        # Show statistics
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Items", len(merged_data["items"]))
                        with col2:
                            total_annotations = sum(len(item.get("annotations", [])) for item in merged_data["items"])
                            st.metric("Total Annotations", total_annotations)
                        with col3:
                            label_count = len(merged_data["categories"]["label"]["labels"])
                            st.metric("Total Labels", label_count)

                        st.divider()

                        # Show label summary
                        st.subheader("🏷️ Merged Labels Summary")
                        label_data = []
                        for label_obj in merged_data["categories"]["label"]["labels"]:
                            label_name = label_obj.get("name", "Unknown")
                            label_data.append({"Label Name": label_name})
                        st.dataframe(pd.DataFrame(label_data), use_container_width=True, hide_index=True)

                        st.divider()

                        # Download button
                        st.subheader("💾 Download Merged JSON")
                        st.download_button(
                            label="📥 Download Merged JSON",
                            data=merged_json,
                            file_name="merged_annotations.json",
                            mime="application/json",
                            use_container_width=True
                        )

                except Exception as e:
                    error_msg = str(e)
                    if "Duplicate label names found" in error_msg:
                        st.error("❌ Cannot merge files with duplicate label names")
                        st.error(error_msg)
                    else:
                        st.error(f"❌ Error merging files: {str(e)}")
                        import traceback
                        st.error(traceback.format_exc())


def main():
    st.set_page_config(page_title="Datumaro Tools", layout="wide")

    st.markdown("""
        <style>
            .main { max-width: 1200px; margin: 0 auto; }
            [data-testid="stAppViewContainer"] { padding-top: 2rem; }
        </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.title("🗂️ Datumaro Tools")
        st.markdown("---")
        st.markdown("**Select a Page:**")
        page = st.radio(
            "page_nav",
            ["📊 Annotation Converter", "🔷 Polygon Point Counter", "🔀 JSON Split & Merge"],
            label_visibility="collapsed"
        )
        st.markdown("---")
        st.markdown("**📊 Annotation Converter**")
        st.caption("Convert Datumaro JSON to Excel with label breakdowns.")
        st.markdown("**🔷 Polygon Point Counter**")
        st.caption("Analyze per-polygon vertex (point) counts.")
        st.markdown("**🔀 JSON Split & Merge**")
        st.caption("Filter annotations by label or merge multiple JSON files.")

    if page == "📊 Annotation Converter":
        annotation_converter_page()
    elif page == "🔷 Polygon Point Counter":
        polygon_point_counter_page()
    else:
        json_split_merge_page()


if __name__ == "__main__":
    main()
