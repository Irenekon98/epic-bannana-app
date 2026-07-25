import streamlit as st
import cv2
import numpy as np
import os
import rawpy
import openpyxl
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import time

# --- ΡΥΘΜΙΣΗ ΣΕΛΙΔΑΣ ---
st.set_page_config(
    page_title="Banana App Web Edition",
    page_icon="🍌",
    layout="wide"
)

# Στυλ για ζεστό μπανανί φόντο (παρόμοιο με το desktop app)
st.markdown(
    """
    <style>
    .stApp {
        background-color: #FEF9C3;
    }
    h1, h2, h3 {
        color: #D97706;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("🍌 Banana App Pro — Web Edition")
st.markdown("**Advanced CIELAB & Ripeness Analytics** (Προσβάσιμο από κινητό/tablet)")

# --- SIDEBAR: ΕΠΙΛΟΓΕΣ ---
st.sidebar.header("⚙️ Ρυθμίσεις Εξαγωγής")
sample_mode = st.sidebar.selectbox(
    "Καταγραφή Pixels στο Excel:",
    ["1 στα 5 Pixels (20%)", "Όλα τα Pixels (100%)"]
)
sample_step = 5 if "1 στα 5" in sample_mode else 1

show_labels = st.sidebar.checkbox("Εμφάνιση Labels (Slice_1, κτλ) στην εικόνα", value=True)

st.sidebar.markdown("---")
uploaded_file = st.sidebar.file_uploader("📷 Επιλογή Αρχείου .NEF", type=["nef", "NEF"])

# --- ΚΥΡΙΩΣ ΛΟΓΙΚΗ ---
if uploaded_file is not None:
    # Αποθηκεύουμε προσωρινά το αρχείο για να το διαβάσει το rawpy
    temp_filename = "temp_uploaded.nef"
    with open(temp_filename, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    file_name = os.path.splitext(uploaded_file.name)[0]
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📥 Αρχικό Αρχείο")
        st.write(f"**Όνομα:** {uploaded_file.name}")
        
        if st.button("🚀 Έναρξη Ανάλυσης", type="primary"):
            with st.spinner("⏳ Γίνεται ανάπτυξη RAW και υπολογισμός CIELAB..."):
                try:
                    # 1. Ανάπτυξη RAW
                    with rawpy.imread(temp_filename) as raw:
                        img_rgb = raw.postprocess(use_camera_wb=True) 
                    
                    img_bgr = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)
                    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
                    blurred = cv2.GaussianBlur(gray, (5, 5), 0)

                    # 2. Otsu Threshold & Mask
                    thresh_val, banana_mask = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                    
                    mask_floodfill = banana_mask.copy()
                    h_m, w_m = banana_mask.shape[:2]
                    floodfill_mask = np.zeros((h_m + 2, w_m + 2), np.uint8)
                    cv2.floodFill(mask_floodfill, floodfill_mask, (0,0), 255)
                    banana_mask = banana_mask | cv2.bitwise_not(mask_floodfill)

                    num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(banana_mask)
                    valid_objects = [(i, centroids[i], stats[i, cv2.CC_STAT_AREA]) for i in range(1, num_labels) if stats[i, cv2.CC_STAT_AREA] >= 5000]

                    if len(valid_objects) < 6:
                        st.error(f"❌ Σφάλμα: Βρέθηκαν μόνο {len(valid_objects)} αντικείμενα (απαιτούνται 6).")
                    else:
                        valid_objects.sort(key=lambda x: x[2], reverse=True)
                        valid_objects = valid_objects[:6]
                        valid_objects.sort(key=lambda x: x[1][1]) 
                        ordered_objects = sorted(valid_objects[:3], key=lambda x: x[1][0]) + sorted(valid_objects[3:6], key=lambda x: x[1][0])

                        clean_mask = np.zeros_like(banana_mask)
                        for obj_id, _, _ in ordered_objects: 
                            clean_mask[labels == obj_id] = 255

                        bg_color = np.array([70, 55, 45], dtype=np.uint8)
                        img_bgr_cleaned = np.where(clean_mask[:, :, None] == 255, img_bgr, bg_color)
                        
                        img_float = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB).astype(np.float32) / 255.0
                        img_lab = cv2.cvtColor(img_float, cv2.COLOR_RGB2Lab)

                        output_map = img_bgr_cleaned.copy()
                        slices_data = {}
                        summary_rows = []

                        # 3. Υπολογισμοί CIELAB ανά Slice
                        for idx, (obj_id, (cX, cY), _) in enumerate(ordered_objects):
                            slice_name = f"Slice_{idx + 1}"
                            obj_mask = (labels == obj_id)
                            
                            L_all = img_lab[:, :, 0][obj_mask]
                            a_all = img_lab[:, :, 1][obj_mask]
                            b_all = img_lab[:, :, 2][obj_mask]

                            L_mean, L_std, L_min, L_max = np.mean(L_all), np.std(L_all), np.min(L_all), np.max(L_all)
                            a_mean, a_std = np.mean(a_all), np.std(a_all)
                            b_mean, b_std = np.mean(b_all), np.std(b_all)
                            Chroma_mean = np.sqrt(a_mean**2 + b_mean**2)
                            Hue_mean = np.degrees(np.arctan2(b_mean, a_mean)) % 360

                            summary_rows.append([
                                slice_name, len(L_all), 
                                round(float(L_mean), 2), round(float(L_std), 2), round(float(L_min), 2), round(float(L_max), 2),
                                round(float(a_mean), 2), round(float(a_std), 2),
                                round(float(b_mean), 2), round(float(b_std), 2),
                                round(float(Chroma_mean), 2), round(float(Hue_mean), 2)
                            ])

                            L_sub, a_sub, b_sub = L_all[::sample_step], a_all[::sample_step], b_all[::sample_step]
                            slices_data[slice_name] = np.column_stack((np.arange(1, len(L_sub) + 1), L_sub, a_sub, b_sub))

                            if show_labels:
                                ys, xs = np.where(obj_mask)
                                min_y = np.min(ys)
                                label_x = int(cX) - 130
                                label_y = max(min_y - 40, 50)
                                cv2.putText(output_map, slice_name, (label_x, label_y), cv2.FONT_HERSHEY_SIMPLEX, 2.5, (0, 0, 0), 12)
                                cv2.putText(output_map, slice_name, (label_x, label_y), cv2.FONT_HERSHEY_SIMPLEX, 2.5, (255, 255, 255), 5)

                        # Αποθήκευση εικόνας αποτελέσματος
                        output_img_path = f"{file_name}_visual_map.png"
                        cv2.imwrite(output_img_path, output_map)

                        # Δημιουργία Excel Summary
                        excel_summary_path = f"{file_name}_summary_analysis.xlsx"
                        wb_sum = openpyxl.Workbook()
                        ws_sum = wb_sum.active
                        ws_sum.title = "Summary_Stats"
                        headers_summary = ['Slice', 'Total_Pixels', 'L*_Mean', 'L*_StdDev', 'L*_Min', 'L*_Max', 'a*_Mean', 'a*_StdDev', 'b*_Mean', 'b*_StdDev', 'Chroma_C*', 'Hue_Angle_h°']
                        ws_sum.append(headers_summary)
                        for r in summary_rows: ws_sum.append(r)
                        wb_sum.save(excel_summary_path)

                        st.success("🎉 Η ανάλυση ολοκληρώθηκε με επιτυχία!")
                        
                        # --- ΑΠΟΤΕΛΕΣΜΑΤΑ ΣΤΟ UI ---
                        st.image(output_img_path, caption="Οπτικός Χάρτης Δειγμάτων", use_container_width=True)
                        
                        # Κουμπί λήψης Excel
                        with open(excel_summary_path, "rb") as fp:
                            st.download_button(
                                label="📥 Λήψη Αρχείου Excel (Summary)",
                                data=fp,
                                file_name=excel_summary_path,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                except Exception as e:
                    st.error(f"❌ Προέκυψε σφάλμα κατά την επεξεργασία: {str(e)}")

else:
    st.info("👈 Παρακαλώ ανεβάστε ένα αρχείο .NEF από την πλαϊνή μπάρα για να ξεκινήσετε.")
