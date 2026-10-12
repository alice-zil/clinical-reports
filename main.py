import streamlit as st
import pandas as pd
import zipfile
import base64
import io
import os
import re
from xhtml2pdf import pisa
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

# Set page style
st.set_page_config(page_title="Κλινικές Αναφορές σε PDF", page_icon="📊", layout="centered")

st.title("📊 Αυτόματη Δημιουργία Αναφορών Ασθενών")
st.write(
    "Ανεβάστε το αρχείο Excel/CSV σας. Η εφαρμογή θα εντοπίσει αυτόματα τη γραμμή "
    "εκκίνησης των ασθενών, θα υπολογίσει τις νόρμες/σχόλια, "
    "και θα δημιουργήσει ένα ξεχωριστό PDF για τον καθένα σε ένα αρχείο .zip."
)

# Συναρτήση μετατροπής ελληνικών σε λατινικά (Greeklish) για ονόματα αρχείων
def safe_greeklish(text):
    if not isinstance(text, str):
        text = str(text)
    translation_table = {
        'α': 'a', 'ά': 'a', 'β': 'b', 'γ': 'g', 'δ': 'd', 'ε': 'e', 'έ': 'e',
        'ζ': 'z', 'η': 'i', 'ή': 'i', 'θ': 'th', 'ι': 'i', 'ί': 'i', 'ϊ': 'i', 'ΐ': 'i',
        'κ': 'k', 'λ': 'l', 'μ': 'm', 'ν': 'n', 'ξ': 'x', 'ο': 'o', 'ό': 'o',
        'π': 'p', 'ρ': 'r', 'σ': 's', 'ς': 's', 'τ': 't', 'υ': 'y', 'ύ': 'y', 'ϋ': 'y', 'ΰ': 'y',
        'φ': 'f', 'χ': 'ch', 'ψ': 'ps', 'ω': 'o', 'ώ': 'o',
        'Α': 'A', 'Ά': 'A', 'Β': 'B', 'Γ': 'G', 'Δ': 'D', 'Ε': 'E', 'Έ': 'E',
        'Ζ': 'Z', 'Η': 'I', 'Ή': 'I', 'Θ': 'Th', 'Ι': 'I', 'Ί': 'I', 'Ϊ': 'I',
        'Κ': 'K', 'Λ': 'L', 'Μ': 'M', 'Ν': 'N', 'Ξ': 'X', 'Ο': 'O', 'Ό': 'O',
        'Π': 'P', 'Ρ': 'R', 'Σ': 'S', 'Τ': 'T', 'Υ': 'Y', 'Ύ': 'Y', 'Ϋ': 'Y',
        'Φ': 'F', 'Χ': 'Ch', 'Ψ': 'Ps', 'Ω': 'O', 'Ώ': 'O'
    }
    converted = "".join(translation_table.get(char, char) for char in text)
    clean_text = re.sub(r'[^a-zA-Z0-9_\- ]', '', converted)
    clean_text = clean_text.replace(' ', '_')
    return re.sub(r'_+', '_', clean_text).strip('_')

# Φόρτωση γραμματοσειρών
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REGULAR_FONT_PATH = os.path.join(BASE_DIR, "DejaVuSans.ttf")
BOLD_FONT_PATH = os.path.join(BASE_DIR, "DejaVuSans-Bold.ttf")

if os.path.exists(REGULAR_FONT_PATH) and os.path.exists(BOLD_FONT_PATH):
    try:
        pdfmetrics.registerFont(TTFont('Helvetica', REGULAR_FONT_PATH))
        pdfmetrics.registerFont(TTFont('Helvetica-Bold', BOLD_FONT_PATH))
    except Exception as e:
        st.error(f"Σφάλμα κατά την εγγραφή των γραμματοσειρών: {e}")
        st.stop()
else:
    st.error(f"⚠️ Τα αρχεία 'DejaVuSans.ttf' και 'DejaVuSans-Bold.ttf' πρέπει να βρίσκονται στον φάκελο: {BASE_DIR}")
    st.stop()

# Βοηθητική συνάρτηση για να βρίσκει το xhtml2pdf τα αρχεία απευθείας χωρίς cache bug


# Βρίσκουμε αυτόματα τα αρχεία στον φάκελο (ανεξάρτητα από πεζά/κεφαλαία στο Linux)
def get_image_path(default_name, alt_keywords):
    p = os.path.join(BASE_DIR, default_name)
    if os.path.isfile(p):
        return p
    if os.path.exists(BASE_DIR):
        for f in os.listdir(BASE_DIR):
            f_low = f.lower()
            if f_low == default_name.lower() or any(k.lower() in f_low for k in alt_keywords):
                return os.path.join(BASE_DIR, f)
    return p

def image_to_data_uri(path):
    if not os.path.isfile(path):
        return None
    with open(path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("utf-8")
    return f"data:image/png;base64,{encoded}"

LOGO1_PATH = get_image_path("logo1.png", ["logo1", "logo_left"])
LOGO2_PATH = get_image_path("logo2.png", ["logo2", "comfortage", "logo_right"])

LOGO1_URI = image_to_data_uri(LOGO1_PATH)
LOGO2_URI = image_to_data_uri(LOGO2_PATH)

# Φόρτωση των 2 λογοτύπων (π.χ. logo1.png και logo2.png στον ίδιο φάκελο)


# --- ΒΟΗΘΗΤΙΚΟΣ ΑΛΓΟΡΙΘΜΟΣ ΑΝΙΧΝΕΥΣΗΣ ΦΥΛΟΥ ΑΠΟ ΤΟ ΟΝΟΜΑ ---
def infer_gender(name):
    name = str(name).strip().upper()
    if not name:
        return 'Female'
    words = name.split()
    s_count = sum(1 for w in words if w.endswith('Σ'))
    if s_count >= 1:
        return 'Male'
    return 'Female'


# --- ΣΥΝΑΡΤΗΣΗ ΑΞΙΟΛΟΓΗΣΗΣ ME ΒΑΣΗ ΤΙΣ ΝΟΡΜΕΣ ---
def evaluate_results(row_data):
    evals = {}
    
    def to_float(val):
        if val is None:
            return None
        val_s = str(val).strip()
        if val_s == '-' or val_s == '' or val_s.lower() == 'nan':
            return None
        try:
            return float(val_s.replace(',', '.'))
        except ValueError:
            return None

    def is_no_test(val):
        if val is None:
            return True
        val_s = str(val).strip()
        return val_s == '-' or val_s == '' or val_s.lower() == 'nan'

    age = to_float(row_data.get('Age'))
    gender = infer_gender(row_data.get('Name'))
    bmi = to_float(row_data.get('BMI'))
    weight = to_float(row_data.get('Weight'))
    height = to_float(row_data.get('Height'))

    # --- ΥΠΟΛΟΓΙΣΜΟΣ ΥΠΟΓΡΑΜΜΙΣΗΣ ΚΑΤΗΓΟΡΙΑΣ BMI ---
    evals['css_bmi_under'] = ""
    evals['css_bmi_normal'] = ""
    evals['css_bmi_over'] = ""
    evals['css_bmi_obese1'] = ""
    evals['css_bmi_obese2'] = ""
    evals['css_bmi_obese3'] = ""

    highlight_style = 'style="text-decoration: underline; font-family: \'Helvetica-Bold\'; background-color: #fcf3cf;"'

    if bmi is not None:
        if bmi < 18.5:
            evals['css_bmi_under'] = highlight_style
        elif 18.5 <= bmi <= 24.99:
            evals['css_bmi_normal'] = highlight_style
        elif 25.0 <= bmi <= 29.99:
            evals['css_bmi_over'] = highlight_style
        elif 30.0 <= bmi <= 34.99:
            evals['css_bmi_obese1'] = highlight_style
        elif 35.0 <= bmi <= 39.99:
            evals['css_bmi_obese2'] = highlight_style
        else:
            evals['css_bmi_obese3'] = highlight_style

    # --- ΥΠΟΛΟΓΙΣΜΟΣ CHAIR STAND TEST NORMS ---
    reps = to_float(row_data.get('ChairStand'))
    if is_no_test(row_data.get('ChairStand')):
        evals['ChairStand_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['CS_Eval'] = "-"
        evals['CS_Color'] = "#7F8C8D"
        evals['CS_Comment'] = "Δεν πραγματοποιήθηκε"
    elif reps is not None and age is not None:
        norms_table = [
            ((60, 64), {'Male': (14, 19), 'Female': (12, 17)}),
            ((65, 69), {'Male': (12, 18), 'Female': (11, 16)}),
            ((70, 74), {'Male': (12, 17), 'Female': (10, 15)}),
            ((75, 79), {'Male': (11, 17), 'Female': (10, 15)}),
            ((80, 84), {'Male': (10, 15), 'Female': (9, 14)}),
            ((85, 89), {'Male': (8, 14), 'Female': (8, 13)}),
            ((90, 150), {'Male': (7, 12), 'Female': (4, 11)})
        ]
        
        selected_norm = None
        age_label = ""
        for (age_range, gender_norms) in norms_table:
            if age_range[0] <= age <= age_range[1]:
                selected_norm = gender_norms.get(gender)
                age_label = "90+" if age_range[0] == 90 else f"{age_range[0]}-{age_range[1]}"
                break
        
        if selected_norm:
            low_val, high_val = selected_norm
            if reps < low_val:
                relation = '<span style="color: #C0392B; font-family: \'Helvetica-Bold\';">κάτω</span>'
                evals['CS_Eval'] = "Κάτω του μετρίου"
                evals['CS_Color'] = "#C0392B"
            elif reps > high_val:
                relation = '<span style="color: #27AE60; font-family: \'Helvetica-Bold\';">άνω</span>'
                evals['CS_Eval'] = "Άνω του μετρίου"
                evals['CS_Color'] = "#27AE60"
            else:
                relation = '<span style="color: #2ECC71; font-family: \'Helvetica-Bold\';">εντός</span>'
                evals['CS_Eval'] = "Φυσιολογική"
                evals['CS_Color'] = "#2ECC71"
                
            evals['ChairStand_Statement'] = (
                f"Εκτελέσατε <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{int(reps)}</b> επαναλήψεις, "
                f"επίδοση που βρίσκεται {relation} των φυσιολογικών τιμών αναφοράς "
                f"(<span style=\"color: #C0392B; font-family: 'Helvetica-Bold';\">{age_label} ετών: {low_val}-{high_val} επαναλήψεις</span>) "
                f"για την ηλικία και το φύλο σας."
            )
            evals['CS_Comment'] = f"Φυσιολογικό εύρος ηλικιακής ομάδας: {low_val}-{high_val} επαναλήψεις."
        else:
            evals['ChairStand_Statement'] = "Η ηλικία του εξεταζόμενου βρίσκεται εκτός των διαθέσιμων ορίων αναφοράς (κάτω των 60 ετών)."
            evals['CS_Eval'] = "Εκτός Ορίων"
            evals['CS_Color'] = "#7F8C8D"
            evals['CS_Comment'] = "Η ηλικία είναι κάτω από 60 ετών."
    else:
        evals['ChairStand_Statement'] = "Δεν καταγράφηκε έγκυρη επίδοση (επαναλήψεις) ή ηλικία για τον υπολογισμό του Chair Stand Test."
        evals['CS_Eval'] = "Μη διαθέσιμο"
        evals['CS_Color'] = "#7F8C8D"
        evals['CS_Comment'] = "Δεν καταγράφηκε έγκυρη τιμή."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ HANDGRIP STRENGTH NORMS (ΑΝΑ ΧΕΡΙ) ---
    raw_gl = row_data.get('Handgrip_L')
    raw_gr = row_data.get('Handgrip_R')
    if is_no_test(raw_gl) and is_no_test(raw_gr):
        evals['HandGrip_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['Grip_Eval'] = "-"
        evals['Grip_Color'] = "#7F8C8D"
        evals['Grip_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        g_l = to_float(raw_gl)
        g_r = to_float(raw_gr)
        if (g_l is not None or g_r is not None) and age is not None:
            grip_norms = [
                ((10, 11), {'Male': (12.6, 22.4), 'Female': (11.8, 21.6)}),
                ((12, 13), {'Male': (19.4, 31.2), 'Female': (14.6, 24.4)}),
                ((14, 15), {'Male': (28.5, 44.3), 'Female': (15.5, 27.3)}),
                ((16, 17), {'Male': (32.6, 52.4), 'Female': (17.2, 29.0)}),
                ((18, 19), {'Male': (35.7, 55.5), 'Female': (19.2, 31.0)}),
                ((20, 24), {'Male': (36.8, 56.6), 'Female': (21.5, 35.3)}),
                ((25, 29), {'Male': (37.7, 57.5), 'Female': (25.6, 41.4)}),
                ((30, 34), {'Male': (36.0, 55.8), 'Female': (21.5, 35.3)}),
                ((35, 39), {'Male': (35.8, 55.6), 'Female': (20.3, 34.1)}),
                ((40, 44), {'Male': (35.5, 55.3), 'Female': (18.9, 32.7)}),
                ((45, 49), {'Male': (34.7, 54.5), 'Female': (18.6, 32.4)}),
                ((50, 54), {'Male': (32.9, 50.7), 'Female': (18.1, 31.9)}),
                ((55, 59), {'Male': (30.7, 48.5), 'Female': (17.7, 31.5)}),
                ((60, 64), {'Male': (30.2, 48.0), 'Female': (17.2, 31.0)}),
                ((65, 69), {'Male': (28.2, 44.0), 'Female': (15.4, 27.2)}),
                ((70, 99), {'Male': (21.3, 35.1), 'Female': (14.7, 24.5)}),
                ((100, 150), {'Male': (21.3, 35.1), 'Female': (14.7, 24.5)})
            ]
            
            selected_grip_norm = None
            age_label = ""
            for (age_range, gender_norms) in grip_norms:
                if age_range[0] <= age <= age_range[1]:
                    selected_grip_norm = gender_norms.get(gender)
                    age_label = f"{age_range[0]}-{age_range[1]}"
                    break
            
            if selected_grip_norm:
                weak_val, strong_val = selected_grip_norm
                statements = []
                
                # Αξιολόγηση Δεξιού Χεριού
                if g_r is not None:
                    if g_r < weak_val:
                        rel_r = "κάτω των φυσιολογικών ορίων"
                    elif g_r > strong_val:
                        rel_r = "άνω των φυσιολογικών ορίων"
                    else:
                        rel_r = "εντός των φυσιολογικών ορίων για το φύλο και την ηλικία σας"
                    statements.append(f"Η μέγιστη δύναμη λαβής για το δεξί χέρι ήταν <b>{g_r}</b> kg, τιμή που βρίσκεται {rel_r} ({age_label} ετών: {weak_val}-{strong_val} kg).")
                
                # Αξιολόγηση Αριστερού Χεριού
                if g_l is not None:
                    if g_l < weak_val:
                        rel_l = "κάτω των φυσιολογικών ορίων"
                    elif g_l > strong_val:
                        rel_l = "άνω των φυσιολογικών ορίων"
                    else:
                        rel_l = "εντός των φυσιολογικών ορίων για το φύλο και την ηλικία σας"
                    statements.append(f"Η μέγιστη δύναμη λαβής για το αριστερό χέρι ήταν <b>{g_l}</b> kg, τιμή που βρίσκεται {rel_l} ({age_label} ετών: {weak_val}-{strong_val} kg).")
                
                # Τελικό κείμενο παραγράφου
                statements.append("Υψηλότερες τιμές υποδηλώνουν καλύτερη μυϊκή δύναμη, ενώ χαμηλότερες τιμές σχετίζονται με μειωμένη μυϊκή λειτουργία.")
                evals['HandGrip_Statement'] = " ".join(statements)
                
                # Κρατάμε τη μέγιστη τιμή για το status του γενικού πίνακα στο τέλος
                max_grip = max([g for g in [g_l, g_r] if g is not None])
                if max_grip < weak_val:
                    evals['Grip_Eval'] = "Αδύναμη"
                    evals['Grip_Color'] = "#C0392B"
                elif max_grip > strong_val:
                    evals['Grip_Eval'] = "Ισχυρή"
                    evals['Grip_Color'] = "#27AE60"
                else:
                    evals['Grip_Eval'] = "Φυσιολογική"
                    evals['Grip_Color'] = "#2ECC71"
                evals['Grip_Comment'] = f"Εύρος αναφοράς ηλικιακής ομάδας: {weak_val}-{strong_val} kg."
                
            else:
                evals['HandGrip_Statement'] = "Δεν βρέθηκαν διαθέσιμα όρια αναφοράς για την ηλικία του εξεταζόμενου."
                evals['Grip_Eval'] = "Μη διαθέσιμο"
                evals['Grip_Color'] = "#7F8C8D"
                evals['Grip_Comment'] = "Η ηλικία είναι εκτός ορίων."
        else:
            evals['HandGrip_Statement'] = "Δεν καταγράφηκε έγκυρη τιμή δυναμομέτρησης ή ηλικίας."
            evals['Grip_Eval'] = "Μη διαθέσιμο"
            evals['Grip_Color'] = "#7F8C8D"
            evals['Grip_Comment'] = "Δεν καταγράφηκε έγκυρη τιμή."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ BERG BALANCE SCALE (BBS) NORMS ---
    if is_no_test(row_data.get('BergBalance')):
        evals['BBS_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['BBS_Eval'] = "-"
        evals['BBS_Color'] = "#7F8C8D"
        evals['BBS_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        bbs = to_float(row_data.get('BergBalance'))
        if bbs is not None:
            if bbs >= 41:
                rating = "Χαμηλός Κίνδυνος"
                color = "#27AE60"
                relation_text = '<span style="color: #27AE60; font-family: \'Helvetica-Bold\';">χαμηλό</span>'
                comment = "Η στατική και δυναμική ισορροπία κρίνεται ικανοποιητική."
            elif bbs >= 21:
                rating = "Μέτριος Κίνδυνος"
                color = "#F39C12"
                relation_text = '<span style="color: #F39C12; font-family: \'Helvetica-Bold\';">μέτριο</span>'
                comment = "Υπάρχει ήπια αστάθεια. Προτείνεται εστιασμένο πρόγραμμα ενδυνάμωσης."
            else:
                rating = "Υψηλός Κίνδυνος"
                color = "#C0392B"
                relation_text = '<span style="color: #C0392B; font-family: \'Helvetica-Bold\';">υψηλό</span>'
                comment = "Σημαντική αστάθεια. Συνιστάται στενή επίβλεψη."
                
            evals['BBS_Statement'] = (
                f"Συγκεντρώσατε <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{int(bbs)}</b>/56 βαθμούς, "
                f"αποτέλεσμα που αντιστοιχεί σε {relation_text} κίνδυνο πτώσης."
            )
            evals['BBS_Eval'] = rating
            evals['BBS_Color'] = color
            evals['BBS_Comment'] = comment
        else:
            evals['BBS_Statement'] = "Δεν καταγράφηκε έγκυρη βαθμολογία για το Berg Balance Scale."
            evals['BBS_Eval'] = "Μη διαθέσιμο"
            evals['BBS_Color'] = "#7F8C8D"
            evals['BBS_Comment'] = "Δεν καταγράφηκε τιμή."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ TINETTI NORMS ---
    if is_no_test(row_data.get('Tinetti')):
        evals['Tinetti_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['Tinetti_Eval'] = "-"
        evals['Tinetti_Color'] = "#7F8C8D"
        evals['Tinetti_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        tinetti = to_float(row_data.get('Tinetti'))
        if tinetti is not None:
            if tinetti >= 25:
                rating = "Χαμηλός Κίνδυνος"
                color = "#27AE60"
                relation_text = '<span style="color: #27AE60; font-family: \'Helvetica-Bold\';">χαμηλό</span>'
                comment = "Πολύ καλή ανταπόκριση στις δοκιμασίες στατικής και δυναμικής ισορροπίας."
            elif tinetti >= 19:
                rating = "Μέτριος Κίνδυνος"
                color = "#F39C12"
                relation_text = '<span style="color: #F39C12; font-family: \'Helvetica-Bold\';">μέτριο</span>'
                comment = "Μέτρια κινητικότητα. Χρήζει ενίσχυσης μέσω στοχευμένης φυσικοθεραπείας."
            else:
                rating = "Υψηλός Κίνδυνος"
                color = "#C0392B"
                relation_text = '<span style="color: #C0392B; font-family: \'Helvetica-Bold\';">υψηλό</span>'
                comment = "Σημαντικό έλλειμμα σε ισορροπία και βάδιση. Απαιτείται αυξημένη προσοχή."
                
            evals['Tinetti_Statement'] = (
                f"Συγκεντρώσατε <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{int(tinetti)}</b>/28 βαθμούς, "
                f"αποτέλεσμα που αντιστοιχεί σε {relation_text} κίνδυνο πτώσης."
            )
            evals['Tinetti_Eval'] = rating
            evals['Tinetti_Color'] = color
            evals['Tinetti_Comment'] = comment
        else:
            evals['Tinetti_Statement'] = "Δεν καταγράφηκε έγκυρη βαθμολογία για το Tinetti Test."
            evals['Tinetti_Eval'] = "Μη διαθέσιμο"
            evals['Tinetti_Color'] = "#7F8C8D"
            evals['Tinetti_Comment'] = "Δεν καταγράφηκε τιμή."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ 6-MINUTE WALK TEST (6MWT) NORMS ---
    if is_no_test(row_data.get('SixMinWalk')):
        evals['Walk_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['Walk_Eval'] = "-"
        evals['Walk_Color'] = "#7F8C8D"
        evals['Walk_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        walk_dist = to_float(row_data.get('SixMinWalk'))
        if walk_dist is not None and age is not None and weight is not None and height is not None:
            if gender == 'Male':
                predicted_dist = (7.57 * height) - (5.02 * age) - (1.76 * weight) - 309
            else:
                predicted_dist = (2.11 * height) - (2.29 * age) - (5.78 * weight) + 667
                
            evals['Walk_Statement'] = (
                f"Διανύσατε μια συνολική απόσταση <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{int(walk_dist)}</b> μέτρων, "
                f"η αναμενόμενη απόσταση βάσει των χαρακτηριστικών σας είναι <b style=\"font-family: 'Helvetica-Bold'; color: #1a5276;\">{predicted_dist:.1f} μέτρα (Enright & Sherill (1988))</b>."
            )
            evals['Walk_Eval'] = f"{int(walk_dist)} μέτρα"
            evals['Walk_Color'] = "#1a5276"
            evals['Walk_Comment'] = f"Αναμενόμενη απόσταση βάσει εξίσωσης: {predicted_dist:.1f} μέτρα."
        else:
            evals['Walk_Statement'] = "Δεν καταγράφηκαν επαρκή δεδομένα (απόσταση, βάρος, ύψος) για τον υπολογισμό του 6MWT."
            evals['Walk_Eval'] = "Μη διαθέσιμο"
            evals['Walk_Color'] = "#7F8C8D"
            evals['Walk_Comment'] = "Ελλιπή στοιχεία σωματομετρίας."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ SINGLE LEG STANCE NORMS ---
    raw_sll = row_data.get('SingleLeg_L')
    raw_slr = row_data.get('SingleLeg_R')
    if is_no_test(raw_sll) and is_no_test(raw_slr):
        evals['SLS_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['SLS_Eval'] = "-"
        evals['SLS_Color'] = "#7F8C8D"
        evals['SLS_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        sl_l = to_float(raw_sll)
        sl_r = to_float(raw_slr)
        if sl_l is not None and sl_r is not None:
            def get_sls_rating(sec):
                if sec > 50: return "εξαιρετική"
                elif 40 <= sec <= 50: return "καλή"
                elif 25 <= sec <= 39: return "μέτρια (μέσος όρος)"
                elif 10 <= sec <= 24: return "μέτρια"
                else: return "φτωχή"
                
            rating_r = get_sls_rating(sl_r)
            rating_l = get_sls_rating(sl_l)
            
            evals['SLS_Statement'] = (
                f"Διατηρήσατε τη θέση μονοποδικής στήριξης για <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{sl_r}</b> δευτερόλεπτα στο δεξί πόδι "
                f"και για <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{sl_l}</b> δευτερόλεπτα στο αριστερό πόδι, αποτέλεσμα που υποδηλώνει "
                f"<span style=\"font-family: 'Helvetica-Bold';\">{rating_r}</span> στατική ισορροπία του δεξιού ποδιού και "
                f"<span style=\"font-family: 'Helvetica-Bold';\">{rating_l}</span> στατική ισορροπία στο αριστερό πόδι."
            )
            
            if sl_r <= 24 or sl_l <= 24:
                evals['SLS_Eval'] = "Μειωμένη"
                evals['SLS_Color'] = "#C0392B"
            else:
                evals['SLS_Eval'] = "Ικανοποιητική"
                evals['SLS_Color'] = "#27AE60"
            evals['SLS_Comment'] = f"Δεξί: {rating_r} | Αριστερό: {rating_l}."
        else:
            evals['SLS_Statement'] = "Δεν καταγράφηκαν έγκυρες μετρήσεις χρόνου για τη δοκιμασία Single Leg Stance."
            evals['SLS_Eval'] = "Μη διαθέσιμο"
            evals['SLS_Color'] = "#7F8C8D"
            evals['SLS_Comment'] = "Ελλιπή δεδομένα."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ 3-METER BACKWARDS WALK TEST NORMS ---
    raw_bw1 = row_data.get('ThreeMBackwards_T1')
    raw_bw2 = row_data.get('ThreeMBackwards_T2')
    if is_no_test(raw_bw1) and is_no_test(raw_bw2):
        evals['MBWT_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['MBWT_Eval'] = "-"
        evals['MBWT_Color'] = "#7F8C8D"
        evals['MBWT_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        bw_t1 = to_float(raw_bw1)
        bw_t2 = to_float(raw_bw2)
        bw_times = [t for t in [bw_t1, bw_t2] if t is not None]
        
        if bw_times:
            avg_bw = sum(bw_times) / len(bw_times)
            if avg_bw <= 3.0:
                bw_relation = '<span style="color: #27AE60; font-family: \'Helvetica-Bold\';">χαμηλή</span>'
                evals['MBWT_Eval'] = "Χαμηλός Κίνδυνος"
                evals['MBWT_Color'] = "#27AE60"
            else:
                bw_relation = '<span style="color: #C0392B; font-family: \'Helvetica-Bold\';">αυξημένη</span>'
                evals['MBWT_Eval'] = "Αυξημένος Κίνδυνος"
                evals['MBWT_Color'] = "#C0392B"
                
            evals['MBWT_Statement'] = (
                f"Εσείς ολοκληρώσατε τη δοκιμασία με μέσο χρόνο <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{avg_bw:.2f}</b> δευτερόλεπτα, "
                f"αυτό το αποτέλεσμα υποδηλώνει {bw_relation} πιθανότητα εμφάνισης πτώσεων."
            )
            evals['MBWT_Comment'] = f"Μέσος χρόνος οπισθοπορείας: {avg_bw:.2f}s (Όριο: <=3.0s)."
        else:
            evals['MBWT_Statement'] = "Δεν καταγράφηκε έγκυρος χρόνος για τη δοκιμασία 3-Meter Backwards Walk Test."
            evals['MBWT_Eval'] = "Μη διαθέσιμο"
            evals['MBWT_Color'] = "#7F8C8D"
            evals['MBWT_Comment'] = "Δεν καταγράφηκε χρόνος."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ SHORT PHYSICAL PERFORMANCE BATTERY (SPPB) NORMS ---
    if is_no_test(row_data.get('SPPB')):
        evals['SPPB_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['SPPB_Eval'] = "-"
        evals['SPPB_Color'] = "#7F8C8D"
        evals['SPPB_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        sppb = to_float(row_data.get('SPPB'))
        if sppb is not None:
            if sppb >= 10:
                rating = "Υψηλή"
                color = "#27AE60"
                perf_text = "υψηλή"
                risk_text = "χαμηλό κίνδυνο"
                frailty_text = "απουσία"
            elif sppb >= 7:
                rating = "Μέτρια"
                color = "#F39C12"
                perf_text = "μέτρια"
                risk_text = "αυξημένη πιθανότητα"
                frailty_text = "παρουσία"
            elif sppb >= 4:
                rating = "Χαμηλή"
                color = "#D35400"
                perf_text = "χαμηλή"
                risk_text = "αυξημένη πιθανότητα"
                frailty_text = "παρουσία"
            else:
                rating = "Σοβαρά Μειωμένη"
                color = "#C0392B"
                perf_text = "σοβαρά μειωμένη"
                risk_text = "αυξημένη πιθανότητα"
                frailty_text = "παρουσία"
                
            evals['SPPB_Statement'] = (
                f"Συγκεντρώσατε <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{int(sppb)}</b>/12 βαθμούς, "
                f"αποτέλεσμα που υποδηλώνει <span style=\"font-family: 'Helvetica-Bold';\">{perf_text}</span> λειτουργική απόδοση με "
                f"<span style=\"font-family: 'Helvetica-Bold';\">{risk_text}</span> λειτουργικής έκπτωσης, και {frailty_text} ευπάθειας."
            )
            evals['SPPB_Eval'] = rating
            evals['SPPB_Color'] = color
            evals['SPPB_Comment'] = f"Συνολικό σκορ SPPB: {int(sppb)}/12."
        else:
            evals['SPPB_Statement'] = "Δεν καταγράφηκε έγκυρη βαθμολογία για το SPPB Test."
            evals['SPPB_Eval'] = "Μη διαθέσιμο"
            evals['SPPB_Color'] = "#7F8C8D"
            evals['SPPB_Comment'] = "Δεν πραγματοποιήθηκε η μέτρηση."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ 10-METER WALK TEST (10MWT) NORMS ---
    raw_vself = row_data.get('TenMWalk_AvgVel_Self')
    raw_vfast = row_data.get('TenMWalk_AvgVel_Fast')
    if is_no_test(raw_vself) and is_no_test(raw_vfast):
        evals['10MWT_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['10MWT_Eval'] = "-"
        evals['10MWT_Color'] = "#7F8C8D"
        evals['10MWT_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        v_self = to_float(raw_vself)
        v_fast = to_float(raw_vfast)
        if v_self is not None and v_fast is not None and age is not None:
            m_norms = {
                29: {'Male': 1.358, 'Female': 1.341},
                39: {'Male': 1.433, 'Female': 1.337},
                49: {'Male': 1.434, 'Female': 1.390},
                59: {'Male': 1.433, 'Female': 1.313},
                69: {'Male': 1.339, 'Female': 1.241},
                79: {'Male': 1.262, 'Female': 1.132},
                150: {'Male': 0.968, 'Female': 0.943}
            }
            
            ref_norm = None
            for max_a, n_vals in sorted(m_norms.items()):
                if age <= max_a:
                    ref_norm = n_vals.get(gender)
                    break
                    
            if ref_norm:
                rel_comfort = "εντός ή άνω των" if v_self >= ref_norm * 0.95 else "κάτω των"
                
                evals['10MWT_Statement'] = (
                    f"Η αυτοεπιλεγόμενη ταχύτητα βάδισής σας ήταν <b style=\"font-family: 'Helvetica-Bold'; color: #1a5276;\">{v_self:.2f} m/s</b> "
                    f"(υγιής τιμή αναφοράς ηλικιακής ομάδας: {ref_norm:.3f} m/s), η οποία κρίνεται ως <b>{rel_comfort}</b> φυσιολογικών ορίων."
                )
                
                if v_self >= ref_norm * 0.95:
                    evals['10MWT_Eval'] = "Φυσιολογική"
                    evals['10MWT_Color'] = "#27AE60"
                else:
                    evals['10MWT_Eval'] = "Μειωμένη"
                    evals['10MWT_Color'] = "#C0392B"
                evals['10MWT_Comment'] = f"Ταχύτητα Άνετης Βάδισης: {v_self:.2f} m/s έναντι νόρμας {ref_norm:.3f} m/s."
            else:
                evals['10MWT_Statement'] = f"Καταγράφηκε αυτοεπιλεγόμενη ταχύτητα {v_self:.2f} m/s."
                evals['10MWT_Eval'] = "Εντός Ορίων"
                evals['10MWT_Color'] = "#2ECC71"
                evals['10MWT_Comment'] = "Δεν προσδιορίστηκε ακριβής νόρμα."
        else:
            evals['10MWT_Statement'] = "Δεν καταγράφηκαν επαρκή δεδομένα ταχύτητας βάδισης (m/s) για το 10MWT."
            evals['10MWT_Eval'] = "Μη διαθέσιμο"
            evals['10MWT_Color'] = "#7F8C8D"
            evals['10MWT_Comment'] = "Ελλιπείς μετρήσεις ταχύτητας."

    # --- ΥΠΟΛΟΓΙΣΜΟΣ TIMED UP AND GO (TUG) NORMS ---
    raw_tug1 = row_data.get('TUG_T1')
    raw_tug2 = row_data.get('TUG_T2')
    if is_no_test(raw_tug1) and is_no_test(raw_tug2):
        evals['TUG_Statement'] = "Δεν πραγματοποιήθηκε."
        evals['TUG_Eval'] = "-"
        evals['TUG_Color'] = "#7F8C8D"
        evals['TUG_Comment'] = "Δεν πραγματοποιήθηκε"
    else:
        tug1 = to_float(raw_tug1)
        tug2 = to_float(raw_tug2)
        tugs = [t for t in [tug1, tug2] if t is not None]
        
        if tugs and age is not None:
            avg_tug = sum(tugs) / len(tugs)
            
            tug_norms_table = [
                ((60, 64), {'Male': (3.8, 5.6), 'Female': (4.4, 6.0)}),
                ((65, 69), {'Male': (4.3, 5.7), 'Female': (4.8, 6.4)}),
                ((70, 74), {'Male': (4.2, 6.0), 'Female': (4.9, 7.1)}),
                ((75, 79), {'Male': (4.6, 7.2), 'Female': (5.2, 7.4)}),
                ((80, 84), {'Male': (5.2, 7.6), 'Female': (5.7, 8.7)}),
                ((85, 89), {'Male': (5.3, 8.9), 'Female': (6.2, 9.6)}),
                ((90, 150), {'Male': (6.2, 10.0), 'Female': (7.3, 11.5)})
            ]
            
            selected_tug_norm = None
            age_label = ""
            for (age_range, gender_norms) in tug_norms_table:
                if age_range[0] <= age <= age_range[1]:
                    selected_tug_norm = gender_norms.get(gender)
                    age_label = f"{age_range[0]}-{age_range[1]}"
                    break
                    
            if selected_tug_norm:
                low_t, high_t = selected_tug_norm
                if avg_tug < low_t:
                    relation = '<span style="color: #27AE60; font-family: \'Helvetica-Bold\';">άνω</span>'
                    evals['TUG_Eval'] = "Άνω του μετρίου"
                    evals['TUG_Color'] = "#27AE60"
                elif avg_tug > high_t:
                    relation = '<span style="color: #C0392B; font-family: \'Helvetica-Bold\';">κάτω</span>'
                    evals['TUG_Eval'] = "Κάτω του μετρίου"
                    evals['TUG_Color'] = "#C0392B"
                else:
                    relation = '<span style="color: #2ECC71; font-family: \'Helvetica-Bold\';">εντός</span>'
                    evals['TUG_Eval'] = "Φυσιολογική"
                    evals['TUG_Color'] = "#2ECC71"
                    
                evals['TUG_Statement'] = (
                    f"Ολοκληρώσατε τη δοκιμασία σε <b style=\"font-family: 'Helvetica-Bold'; color: #C0392B;\">{avg_tug:.1f}</b> δευτερόλεπτα, "
                    f"επίδοση που βρίσκεται {relation} των φυσιολογικών τιμών αναφοράς "
                    f"(<span style=\"color: #C0392B; font-family: 'Helvetica-Bold';\">{age_label} ετών: {low_t:.1f}-{high_t:.1f} δευτερόλεπτα</span>) "
                    f"για την ηλικία και το φύλο σας."
                )
                evals['TUG_Comment'] = f"Φυσιολογικό εύρος ηλικιακής ομάδας: {low_t:.1f}-{high_t:.1f} δευτερόλεπτα."
            else:
                evals['TUG_Statement'] = "Η ηλικία του εξεταζόμενου βρίσκεται εκτός των διαθέσιμων ορίων αναφοράς (κάτω των 60 ετών)."
                evals['TUG_Eval'] = "Εκτός Ορίων"
                evals['TUG_Color'] = "#7F8C8D"
                evals['TUG_Comment'] = "Η ηλικία είναι κάτω από 60 έτη."
        else:
            evals['TUG_Statement'] = "Δεν καταγράφηκε έγκυρος χρόνος ολοκλήρωσης ή ηλικία για το TUG."
            evals['TUG_Eval'] = "Μη διαθέσιμο"
            evals['TUG_Color'] = "#7F8C8D"
            evals['TUG_Comment'] = "Δεν καταγράφηκε έγκυρη τιμή."

    return evals


# HTML Template
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="el">
<head>
    <meta charset="UTF-8">
    <style>
        @page {
            size: a4;
            margin: 15mm 15mm 15mm 15mm;
        }
        body {
            font-family: 'Helvetica', sans-serif;
            color: #2c3e50;
            line-height: 1.5;
        }
        b, strong, th {
            font-family: 'Helvetica-Bold', sans-serif;
            font-weight: normal;
        }
        .header-table {
            width: 100%;
            border-bottom: 3px solid #1a5276;
            margin-bottom: 20px;
            padding-bottom: 5px;
        }
        .title {
            font-size: 15pt;
            color: #1a5276;
            line-height: 1.3;
            margin: 0;
            padding-bottom: 4px;
            font-family: 'Helvetica-Bold', sans-serif;
        }
        .subtitle {
            font-size: 10pt;
            color: #7f8c8d;
            line-height: 1.2;
            margin: 0;
        }
        .header-meta {
            font-size: 9pt;
            color: #566573;
            line-height: 1.4;
            text-align: right;
            vertical-align: bottom;
        }
        .section-title {
            font-size: 11pt;
            color: #1f618d;
            border-left: 4px solid #1f618d;
            padding-left: 8px;
            margin-top: 20px;
            margin-bottom: 10px;
            font-family: 'Helvetica-Bold', sans-serif;
        }
        .sub-section-title {
            font-size: 10pt;
            color: #2c3e50;
            margin-top: 15px;
            margin-bottom: 5px;
            font-family: 'Helvetica-Bold', sans-serif;
        }
        .demographics-table {
            width: 100%;
            margin-bottom: 15px;
            border-collapse: collapse;
        }
        .demographics-table td {
            padding: 7px 10px;
            font-size: 9.5pt;
            border-bottom: 1px solid #e5e8e8;
            vertical-align: middle;
        }
        .demographics-table .label {
            width: 35%;
            color: #34495e;
            font-family: 'Helvetica-Bold', sans-serif;
            background-color: #f8f9f9;
            border-right: 1px solid #e5e8e8;
        }
        .demographics-table .value {
            width: 65%;
            color: #2c3e50;
            padding-left: 12px;
        }
        .grid-table {
            width: 100%;
            margin-bottom: 15px;
            border-collapse: collapse;
        }
        .grid-table th {
            background-color: #2c3e50;
            color: white;
            font-size: 9pt;
            text-align: left;
            padding: 7px 9px;
            font-family: 'Helvetica-Bold', sans-serif;
        }
        .grid-table td {
            padding: 7px 9px;
            font-size: 9pt;
            border: 1px solid #d6dbdf;
            vertical-align: middle;
        }
        .grid-table tbody tr:nth-child(even) {
            background-color: #fcfcfc;
        }
        .badge {
            font-family: 'Helvetica-Bold', sans-serif;
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 8.5pt;
            text-align: center;
            vertical-align: middle;
        }
        .explanation-text {
            font-size: 9pt;
            color: #2c3e50;
            margin-bottom: 8px;
            line-height: 1.4;
        }
        .explanation-text span.red-bold {
            color: #C0392B;
            font-family: 'Helvetica-Bold', sans-serif;
        }
    </style>
</head>
<body>

    <div class="header">
        <table class="header-table" style="width: 100%; border-bottom: 3px solid #1a5276; margin-bottom: 20px; padding-bottom: 8px;">
            <tr>
                <!-- Αριστερό Logo -->
                <td style="width: 20%; vertical-align: middle; text-align: left;">
                    {Logo1_Tag}
                </td>
                
                <!-- Κεντρικός Τίτλος -->
                <td style="width: 60%; vertical-align: middle; text-align: center;">
                    <div class="title" style="font-size: 14pt; margin: 0; padding-bottom: 2px;">ΚΛΙΝΙΚΗ ΑΝΑΦΟΡΑ ΑΞΙΟΛΟΓΗΣΗΣ</div>
                    <div class="subtitle" style="font-size: 9.5pt; margin: 0; padding-bottom: 4px;">Πρόγραμμα Φυσικής & Γνωστικής Παρέμβασης (ACTIVE)</div>
                    <div style="font-size: 8.5pt; color: #566573;">
                        <b>Ημερομηνία:</b> {Date} &nbsp;|&nbsp; <b>Φάση:</b> {Phase}
                    </div>
                </td>
                
                <!-- Δεξί Logo -->
                <td style="width: 20%; vertical-align: middle; text-align: right;">
                    {Logo2_Tag}
                </td>
            </tr>
        </table>
    </div>

    <div class="section-title">Προσωπικά Στοιχεία & Δημογραφικά</div>
    <table class="demographics-table">
        <tr>
            <td class="label">Ονοματεπώνυμο:</td>
            <td class="value"><b>{Name}</b></td>
        </tr>
        <tr>
            <td class="label">Κωδικός (ID):</td>
            <td class="value"><b>{Code}</b></td>
        </tr>
        <tr>
            <td class="label">Ηλικία:</td>
            <td class="value">{Age} έτη</td>
        </tr>
        <tr>
            <td class="label">Έτη Εκπαίδευσης:</td>
            <td class="value">{EducationYears}</td>
        </tr>
        <tr>
            <td class="label">Βάρος / Ύψος:</td>
            <td class="value">{Weight} kg  |  {Height} cm</td>
        </tr>
        <tr>
            <td class="label">Δείκτης Μάζας Σώματος (ΔΜΣ / BMI):</td>
            <td class="value"><b>{BMI}</b></td>
        </tr>
    </table>

    <div class="section-title">Πίνακας Συνοπτικής Παρουσίασης Τιμών</div>
    
    <div class="sub-section-title">Αποτελέσματα Φυσικών & Κινητικών Δοκιμασιών (Σειρά Βάσει Ανάλυσης)</div>
    <table class="grid-table">
        <thead>
            <tr>
                <th width="60%">Δοκιμασία (Test)</th>
                <th width="40%">Επίδοση / Μέτρηση</th>
            </tr>
        </thead>
        <tbody>
            <tr><td><b>Chair Stand Test</b></td><td>{ChairStand} επαναλήψεις</td></tr>
            <tr><td><b>HandGrip Strength</b></td><td>Αριστερό: {Handgrip_L} kg | Δεξί: {Handgrip_R} kg</td></tr>
            <tr><td><b>Berg Balance Scale (BBS)</b></td><td>{BergBalance} / 56</td></tr>
            <tr><td><b>Tinetti Test</b></td><td>{Tinetti} / 28</td></tr>
            <tr><td><b>6 Minute Walk Test</b></td><td>{SixMinWalk} μέτρα</td></tr>
            <tr><td><b>Single Leg Stance</b></td><td>Αριστερό: {SingleLeg_L}s | Δεξί: {SingleLeg_R}s</td></tr>
            <tr><td><b>3m Backwards Walk Test</b></td><td>1η: {ThreeMBackwards_T1}s | 2η: {ThreeMBackwards_T2}s</td></tr>
            <tr><td><b>SPPB Test</b></td><td>{SPPB} / 12</td></tr>
            <tr><td><b>Timed Up And Go (TUG)</b></td><td>1η: {TUG_T1}s | 2η: {TUG_T2}s</td></tr>
        </tbody>
    </table>

    <div class="sub-section-title">Δοκιμασία Βάδισης 10 Μέτρων (10-m Walk Test)</div>
    <table class="grid-table">
        <thead>
            <tr>
                <th width="40%">Παράμετρος</th>
                <th width="30%">Αυτοεπιλεγόμενη Ταχύτητα</th>
                <th width="30%">Γρήγορη Ταχύτητα</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><b>Μέσος Χρόνος (sec)</b></td>
                <td>{TenMWalk_AvgTime_Self} s</td>
                <td>{TenMWalk_AvgTime_Fast} s</td>
            </tr>
            <tr>
                <td><b>Μέση Ταχύτητα (m/s)</b></td>
                <td>{TenMWalk_AvgVel_Self} m/s</td>
                <td>{TenMWalk_AvgVel_Fast} m/s</td>
            </tr>
        </tbody>
    </table>

    <div class="section-title" style="page-break-before: always;">Αναλυτική Αξιολόγηση & Normative Data</div>
    
    <!-- ΔΜΣ (BMI) -->
    <div class="sub-section-title">Ανάλυση Δείκτη Μάζας Σώματος (ΔΜΣ / BMI)</div>
    <div class="explanation-text">
        <span class="red-bold">Δείκτης Μάζας Σώματος (ΔΜΣ):</span> Χρησιμοποιείται ευρέως ως γενικός δείκτης για το εάν ένα άτομο έχει ένα υγιές σωματικό βάρος για το ύψος του.
    </div>
    <table class="grid-table">
        <thead>
            <tr>
                <th width="30%">Κατηγοριοποίηση βάση ΔΜΣ</th>
                <th width="20%">ΔΜΣ (kg/m2)</th>
                <th width="30%">Κατηγοριοποίηση βάση ΔΜΣ</th>
                <th width="20%">ΔΜΣ (kg/m2)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td {css_bmi_under}>Ελλιποβαρής</td>
                <td {css_bmi_under}>&lt; 18,5</td>
                <td {css_bmi_obese1}>Παχυσαρκία τύπου Ι</td>
                <td {css_bmi_obese1}>30,0 – 34,99</td>
            </tr>
            <tr>
                <td {css_bmi_normal}>Φυσιολογικό βάρος</td>
                <td {css_bmi_normal}>18,5 – 24,99</td>
                <td {css_bmi_obese2}>Παχυσαρκία τύπου ΙΙ</td>
                <td {css_bmi_obese2}>35,0 – 39,99</td>
            </tr>
            <tr>
                <td {css_bmi_over}>Υπέρβαρος</td>
                <td {css_bmi_over}>25 – 29,99</td>
                <td {css_bmi_obese3}>Παχυσαρκία τύπου ΙΙΙ</td>
                <td {css_bmi_obese3}>&ge; 40,0</td>
            </tr>
        </tbody>
    </table>

    <!-- CHAIR STAND TEST -->
    <div class="sub-section-title">Ανάλυση Chair Stand Test</div>
    <div class="explanation-text">
        <span class="red-bold">CHAIR STAND TEST:</span> Αξιολογεί τη λειτουργική δύναμη και αντοχή των μυών των κάτω άκρων μέσω του αριθμού των πλήρων επαναλήψεων έγερσης από καθιστή θέση που εκτελεί ο εξεταζόμενος σε 30 δευτερόλεπτα.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {ChairStand_Statement}
    </div>

    <!-- HANDGRIP STRENGTH -->
    <div class="sub-section-title">Ανάλυση HandGrip Strength</div>
    <div class="explanation-text">
        <span class="red-bold">HANDGRIP STRENGTH (Kgr):</span> Μετρά τη μέγιστη ισομετρική δύναμη των μυών του χεριού και του αντιβραχίου.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {HandGrip_Statement}
    </div>

    <!-- BERG BALANCE SCALE (BBS) -->
    <div class="sub-section-title">Ανάλυση Berg Balance Scale (BBS)</div>
    <div class="explanation-text">
        <span class="red-bold">BERG BALANCE SCALE (BBS):</span> Χρησιμοποιείται για την εκτίμηση του κινδύνου πτώσης. Οι υψηλές τιμές υποδηλώνουν καλύτερη λειτουργική ισορροπία και μικρότερο κίνδυνο πτώσης.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {BBS_Statement}
    </div>

    <!-- TINETTI -->
    <div class="sub-section-title">Ανάλυση Tinetti</div>
    <div class="explanation-text">
        <span class="red-bold">TINETTI:</span> Χρησιμοποιείται για την εκτίμηση του κινδύνου πτώσης λαμβάνοντας υπόψη την ισορροπία και την κινητικότητα. Οι υψηλές τιμές υποδηλώνουν καλύτερη ισορροπία και κινητικότητα, καθώς και μικρότερο κίνδυνο πτώσης.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {Tinetti_Statement}
    </div>

    <!-- 6-MINUTE WALK TEST (6MWT) -->
    <div class="sub-section-title">Ανάλυση 6-Minute Walk Test (6MWT)</div>
    <div class="explanation-text">
        <span class="red-bold">6-MINUTE WALK TEST (6MWT):</span> Αξιολογεί τη λειτουργική ικανότητα βάδισης και τη συνολική αντοχή κατά την άσκηση, μέσω της μέγιστης απόστασης που μπορεί να διανύσει ο εξεταζόμενος σε χρονικό διάστημα 6 λεπτών.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {Walk_Statement}
    </div>

    <!-- SINGLE LEG STANCE -->
    <div class="sub-section-title">Ανάλυση Single Leg Stance</div>
    <div class="explanation-text">
        <span class="red-bold">SINGLE LEG STANCE:</span> Αξιολογεί τη στατική ισορροπία χρονομετρώντας για πόσο χρόνο μπορεί ένα άτομο να σταθεί χωρίς βοήθεια στο ένα πόδι.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {SLS_Statement}
    </div>

    <!-- 3-METER BACKWARDS WALK TEST -->
    <div class="sub-section-title">Ανάλυση 3-Meter Backwards Walk Test (MBWT)</div>
    <div class="explanation-text">
        <span class="red-bold">3-METER BACKWARDS WALK TEST (MBWT):</span> Αξιολογεί τη δυναμική ισορροπία και την ικανότητα βάδισης προς τα πίσω μέσω του χρόνου που απαιτείται για την κάλυψη απόστασης 3 μέτρων. Ο μικρότερος χρόνος ολοκλήρωσης υποδηλώνει καλύτερο κινητικό έλεγχο και χαμηλότερο κίνδυνο πτώσης, ενώ ο μεγαλύτερος χρόνος σχετίζεται με αυξημένη πιθανότητα εμφάνισης πτώσεων. Η ερμηνεία του αποτελέσματος είναι η ακόλουθη:
        <br>&bull; <b>&le;3,0 δευτερόλεπτα:</b> Χαμηλή πιθανότητα εμφάνισης πτώσεων.
        <br>&bull; <b>&ge;4,5 δευτερόλεπτα:</b> Αυξημένη πιθανότητα εμφάνισης πτώσεων.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {MBWT_Statement}
    </div>

    <!-- SHORT PHYSICAL PERFORMANCE BATTERY (SPPB) -->
    <div class="sub-section-title">Ανάλυση Short Physical Performance Battery (SPPB)</div>
    <div class="explanation-text">
        <span class="red-bold">SHORT PHYSICAL PERFORMANCE BATTERY (SPPB):</span> Αξιολογεί τη λειτουργική απόδοση των κάτω άκρων μέσω δοκιμασιών στατικής ισορροπίας, ταχύτητας βάδισης και επαναλαμβανόμενης έγερσης από καθιστή θέση.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {SPPB_Statement}
    </div>

    <!-- 10-METER WALK TEST (10MWT) -->
    <div class="sub-section-title">Ανάλυση 10-Meter Walk Test (10MWT)</div>
    <div class="explanation-text">
        <span class="red-bold">10-METER WALK TEST (10MWT):</span> Αξιολογεί τη λειτουργική κινητικότητα, το βάδισμα και την ταχύτητα βάδισης σε μικρή απόσταση 10 μέτρων. Η ερμηνεία βασίζεται κυρίως στην ταχύτητα βάδισης (m/s).
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {10MWT_Statement}
    </div>

    <!-- TIMED UP AND GO (TUG) -->
    <div class="sub-section-title">Ανάλυση Timed Up And Go (TUG)</div>
    <div class="explanation-text">
        <span class="red-bold">TIMED UP AND GO (TUG):</span> Αξιολογεί την ταχύτητα, την ευκινησία και την ισορροπία κατά τη βάδιση μέσω του χρόνου που απαιτείται για την ολοκλήρωση αυτής της δοκιμασίας.
    </div>
    <div class="explanation-text" style="background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; border-radius: 4px; margin-top: 5px;">
        {TUG_Statement}
    </div>

    <!-- ΕΥΘΥΓΡΑΜΜΙΣΜΕΣΟΣ ΠΙΝΑΚΑΣ ΣΥΝΟΨΗΣ ΜΕ ΟΛΑ ΤΑ TEST -->
    <div class="sub-section-title" style="page-break-before: always;">Κλινικός Σχολιασμός & Αξιολόγηση με Νόρμες</div>
    <table class="grid-table">
        <thead>
            <tr>
                <th width="25%">Δοκιμασία</th>
                <th width="25%">Κατάσταση / Νόρμα</th>
                <th width="50%">Κλινική Παρατήρηση</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><b>Chair Stand Test</b></td>
                <td><span class="badge" style="background-color: {CS_Color};">{CS_Eval}</span></td>
                <td>{CS_Comment}</td>
            </tr>
            <tr>
                <td><b>HandGrip (Max)</b></td>
                <td><span class="badge" style="background-color: {Grip_Color};">{Grip_Eval}</span></td>
                <td>{Grip_Comment}</td>
            </tr>
            <tr>
                <td><b>Berg Balance Scale</b></td>
                <td><span class="badge" style="background-color: {BBS_Color};">{BBS_Eval}</span></td>
                <td>{BBS_Comment}</td>
            </tr>
            <tr>
                <td><b>Tinetti Test</b></td>
                <td><span class="badge" style="background-color: {Tinetti_Color};">{Tinetti_Eval}</span></td>
                <td>{Tinetti_Comment}</td>
            </tr>
            <tr>
                <td><b>6 Min Walk (6MWT)</b></td>
                <td><span class="badge" style="background-color: {Walk_Color};">{Walk_Eval}</span></td>
                <td>{Walk_Comment}</td>
            </tr>
            <tr>
                <td><b>Single Leg Stance</b></td>
                <td><span class="badge" style="background-color: {SLS_Color};">{SLS_Eval}</span></td>
                <td>{SLS_Comment}</td>
            </tr>
            <tr>
                <td><b>3m Backwards Test</b></td>
                <td><span class="badge" style="background-color: {MBWT_Color};">{MBWT_Eval}</span></td>
                <td>{MBWT_Comment}</td>
            </tr>
            <tr>
                <td><b>SPPB Test</b></td>
                <td><span class="badge" style="background-color: {SPPB_Color};">{SPPB_Eval}</span></td>
                <td>{SPPB_Comment}</td>
            </tr>
            <tr>
                <td><b>10-Meter Walk Test</b></td>
                <td><span class="badge" style="background-color: {10MWT_Color};">{10MWT_Eval}</span></td>
                <td>{10MWT_Comment}</td>
            </tr>
            <tr>
                <td><b>Timed Up And Go (Avg)</b></td>
                <td><span class="badge" style="background-color: {TUG_Color};">{TUG_Eval}</span></td>
                <td>{TUG_Comment}</td>
            </tr>
        </tbody>
    </table>

</body>
</html>
"""

uploaded_file = st.file_uploader("Επιλέξτε το αρχείο σας (CSV ή Excel)", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(uploaded_file, header=None, encoding='utf-8-sig')
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, encoding='cp1253')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        start_row_idx = None
        for idx, row in df_raw.iterrows():
            first_val = str(row.iloc[0]).strip()
            if first_val in ['1', '1.0', '1']:
                try:
                    if float(first_val) == 1.0:
                        start_row_idx = idx
                        break
                except ValueError:
                    pass

        if start_row_idx is None:
            st.error("⚠️ Δεν βρέθηκε γραμμή που να ξεκινάει με τον αριθμό 1 (Αύξων Αριθμός) στην πρώτη στήλη.")
            st.stop()
        
        cleaned_df = df_raw.iloc[start_row_idx:].copy()
        cleaned_df = cleaned_df.reset_index(drop=True)  # <--- ΠΡΟΣΘΕΣΕ ΑΥΤΗ ΤΗ ΓΡΑΜΜΗ ΕΔΩ
        
        col_names = [
            "AA", "Name", "Code", "Phase", "Date", "Age", "EducationYears", "Weight", "Height", "BMI",
            "ChairStand", "SixMinWalk", "BergBalance", "Tinetti", "SPPB",
            "ThreeMBackwards_T1", "ThreeMBackwards_T2",
            "TUG_T1", "TUG_T2",
            "SingleLeg_L", "SingleLeg_R",
            "Handgrip_L", "Handgrip_R",
            "TenMWalk_AvgTime_Self", "TenMWalk_AvgTime_Fast",
            "TenMWalk_AvgVel_Self", "TenMWalk_AvgVel_Fast"
        ]
        
        if df_raw.shape[1] >= len(col_names):
            cleaned_df = cleaned_df.iloc[:, :len(col_names)]
            cleaned_df.columns = col_names
        else:
            st.error(f"Το αρχείο δεν έχει τον αναμενόμενο αριθμό στηλών ({len(col_names)} στήλες).")
            st.stop()
            
        cleaned_df["AA_numeric"] = pd.to_numeric(cleaned_df["AA"], errors='coerce')
        cleaned_df = cleaned_df[cleaned_df["AA_numeric"].notna()].copy()
        cleaned_df["AA"] = cleaned_df["AA_numeric"].astype(int)
        cleaned_df = cleaned_df.drop(columns=["AA_numeric"])

        st.success(f"📊 Εντοπίστηκε η αρχή των δεδομένων στη γραμμή {start_row_idx + 1}!")
        st.success(f"👥 Εντοπίστηκαν {len(cleaned_df)} ασθενείς προς επεξεργασία!")

        # --- ΜΕΝΟΥ ΕΠΙΛΟΓΗΣ ΑΣΘΕΝΩΝ ΜΕΣΩ ST.FORM ΓΙΑ ΑΜΕΣΗ ΑΝΤΑΠΟΚΡΙΣΗ ---
        # --- ΜΕΝΟΥ ΕΠΙΛΟΓΗΣ ΑΣΘΕΝΩΝ ΧΩΡΙΣ ΦΟΡΜΑ ΓΙΑ ΑΜΕΣΗ ΑΝΤΑΠΟΚΡΙΣΗ ---
        # --- ΜΕΝΟΥ ΕΠΙΛΟΓΗΣ ΑΣΘΕΝΩΝ ---
        st.subheader("👥 Επιλογή Ασθενών για Εξαγωγή")
        
        # Αρχικοποίηση μόνιμου state με βάση το 'AA' (1 έως 76)
        if "selected_patients" not in st.session_state:
            st.session_state.selected_patients = {row["AA"]: True for _, row in cleaned_df.iterrows()}

        for _, row in cleaned_df.iterrows():
            if row["AA"] not in st.session_state.selected_patients:
                st.session_state.selected_patients[row["AA"]] = True

        # Κουμπιά γρήγορης καθολικής αλλαγής
        col_btn1, col_btn2, _ = st.columns([1.5, 1.5, 4])
        if col_btn1.button("✅ Επιλογή Όλων"):
            for aa in st.session_state.selected_patients:
                st.session_state.selected_patients[aa] = True
            if "patient_editor" in st.session_state:
                del st.session_state["patient_editor"]
            st.rerun()
        if col_btn2.button("❌ Αποεπιλογή Όλων"):
            for aa in st.session_state.selected_patients:
                st.session_state.selected_patients[aa] = False
            if "patient_editor" in st.session_state:
                del st.session_state["patient_editor"]
            st.rerun()

        display_df = cleaned_df[["AA", "Name", "Code", "Phase"]].copy()
        display_df["Εξαγωγή"] = display_df["AA"].map(lambda aa: st.session_state.selected_patients.get(aa, True))

        # Εμφάνιση του πίνακα
        edited_df = st.data_editor(
            display_df,
            column_config={
                "Εξαγωγή": st.column_config.CheckboxColumn(
                    "Εξαγωγή PDF;",
                    help="Τσεκάρετε ποιοι ασθενείς θα συμπεριληφθούν",
                    default=True,
                )
            },
            disabled=["AA", "Name", "Code", "Phase"],
            hide_index=True,
            width=700,
            key="patient_editor"
        )
        
        # Ενημέρωση του session_state με βάση το AA
        for _, row in edited_df.iterrows():
            st.session_state.selected_patients[row["AA"]] = bool(row["Εξαγωγή"])

        # Φιλτράρισμα τελικού dataframe απευθείας από τα τσεκαρισμένα κουτάκια
        final_processing_df = cleaned_df[edited_df["Εξαγωγή"].astype(bool).values].copy()

        st.info(f"🎯 Έχουν επιλεγεί **{len(final_processing_df)}** από τους **{len(cleaned_df)}** ασθενείς.")
        
        # --- ΔΗΜΙΟΥΡΓΙΑ PDF ---
                if st.button("🚀 Μετατροπή & Δημιουργία PDF"):
            if len(final_processing_df) == 0:
                st.error("⚠️ Πρέπει να επιλέξετε τουλάχιστον έναν ασθενή.")
                st.stop()

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                progress_bar = st.progress(0)

                for index, (idx, row) in enumerate(final_processing_df.iterrows()):
                    row_dict = row.to_dict()
                    evaluations = evaluate_results(row_dict)

                    row_data = {}
                    for k, v in row_dict.items():
                        if pd.isna(v):
                            row_data[k] = "-"
                        else:
                            val_str = str(v).strip()
                            if k == "Date" and " " in val_str:
                                val_str = val_str.split(" ")[0]
                            row_data[k] = val_str
                    row_data.update(evaluations)

                    # Λογότυπα ως base64 (αριστερά logo1, δεξιά logo2)
                    row_data['Logo1_Tag'] = f'<img src="{LOGO1_URI}" height="55">' if LOGO1_URI else ''
                    row_data['Logo2_Tag'] = f'<img src="{LOGO2_URI}" height="55">' if LOGO2_URI else ''

                    for k, v in row_data.items():
                        if k not in evaluations and k not in ['Logo1_Tag', 'Logo2_Tag'] and str(v).replace('.', '', 1).isdigit():
                            try:
                                row_data[k] = f"{float(v):.2f}".rstrip('0').rstrip('.')
                            except ValueError:
                                pass

                    html_content = HTML_TEMPLATE
                    for k, v in row_data.items():
                        html_content = html_content.replace("{" + k + "}", str(v))

                    pdf_io = io.BytesIO()
                    pisa.CreatePDF(html_content, dest=pdf_io, encoding='utf-8')
                    pdf_bytes = pdf_io.getvalue()

                    latin_name = safe_greeklish(row_data['Name'])
                    latin_phase = safe_greeklish(row_data['Phase'])
                    file_name = f"Reports/{latin_name}_{latin_phase}.pdf"

                    zip_file.writestr(file_name, pdf_bytes)
                    progress_bar.progress((index + 1) / len(final_processing_df))

            st.success("✅ Όλα τα PDF δημιουργήθηκαν με επιτυχία!")

            st.download_button(
                label="📥 Λήψη όλων των PDF (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Patient_Reports.zip",
                mime="application/zip"
            )
            if len(final_processing_df) == 0:
                st.error("⚠️ Πρέπει να επιλέξετε τουλάχιστον έναν ασθενή και να πατήσετε 'Αποθήκευση Επιλογών' για να προχωρήσει η εξαγωγή.")
                st.stop()

            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                progress_bar = st.progress(0)
                
                for index, (idx, row) in enumerate(final_processing_df.iterrows()):
                    row_dict = row.to_dict()
                    
                    evaluations = evaluate_results(row_dict)
                    
                    row_data = {}
                    for k, v in row_dict.items():
                        if pd.isna(v):
                            row_data[k] = "-"
                        else:
                            val_str = str(v).strip()
                            if k == "Date" and " " in val_str:
                                val_str = val_str.split(" ")[0]
                            row_data[k] = val_str
                    row_data.update(evaluations)

                    # Λογότυπα
                    row_data['Logo1_Tag'] = '<img src="logo1.png" height="55">' if os.path.exists(os.path.join(BASE_DIR, "logo1.png")) else ''
                    row_data['Logo2_Tag'] = '<img src="logo2.png" height="55">' if os.path.exists(os.path.join(BASE_DIR, "logo2.png")) else ''
                    
                    for k, v in row_data.items():
                        if k not in evaluations and k not in ['Logo1_Tag', 'Logo2_Tag'] and str(v).replace('.', '', 1).isdigit():
                            try:
                                row_data[k] = f"{float(v):.2f}".rstrip('0').rstrip('.')
                            except ValueError:
                                pass
                    
                    # Δημιουργία περιεχομένου HTML
                    html_content = HTML_TEMPLATE
                    for k, v in row_data.items():
                        html_content = html_content.replace("{" + k + "}", str(v))
                    
                    pdf_io = io.BytesIO()
                    pisa.CreatePDF(html_content, dest=pdf_io, encoding='utf-8', path=BASE_DIR)
                    pdf_bytes = pdf_io.getvalue()
                    
                    latin_name = safe_greeklish(row_data['Name'])
                    latin_phase = safe_greeklish(row_data['Phase'])
                    
                    file_name = f"Reports/{latin_name}_{latin_phase}.pdf"
                    
                    zip_file.writestr(file_name, pdf_bytes)
                    progress_bar.progress((index + 1) / len(final_processing_df))
                    row_dict = row.to_dict()
                    
                    evaluations = evaluate_results(row_dict)
                    
                    row_data = {}
                    for k, v in row_dict.items():
                        if pd.isna(v):
                            row_data[k] = "-"
                        else:
                            val_str = str(v).strip()
                            if k == "Date" and " " in val_str:
                                val_str = val_str.split(" ")[0]
                            row_data[k] = val_str
                    row_data.update(evaluations)

                    # Ανάγνωση των δύο ξεχωριστών αρχείων (όχι base64)
                    row_data['Logo1_Tag'] = '<img src="logo1.png" height="55">' if os.path.exists(os.path.join(BASE_DIR, "logo1.png")) else ''
                    row_data['Logo2_Tag'] = '<img src="logo2.png" height="55">' if os.path.exists(os.path.join(BASE_DIR, "logo2.png")) else ''
                    
                    for k, v in row_data.items():
                        if k not in evaluations and k not in ['Logo1_Tag', 'Logo2_Tag'] and str(v).replace('.', '', 1).isdigit():
                            try:
                                row_data[k] = f"{float(v):.2f}".rstrip('0').rstrip('.')
                            except ValueError:
                                pass
                    
                    # --- ΕΔΩ ΟΡΙΖΕΤΑΙ ΤΟ html_content (ΠΟΥ ΕΙΧΕ ΣΒΗΣΤΕΙ) ---
                    html_content = HTML_TEMPLATE
                    for k, v in row_data.items():
                        html_content = html_content.replace("{" + k + "}", str(v))
                    
                    pdf_io = io.BytesIO()
                    # Χρήση path=BASE_DIR ώστε να διαβάζει τα logo1.png και logo2.png ως διαφορετικά αρχεία
                    pisa.CreatePDF(html_content, dest=pdf_io, encoding='utf-8', path=BASE_DIR)
                    pdf_bytes = pdf_io.getvalue()
                    
                    latin_name = safe_greeklish(row_data['Name'])
                    latin_phase = safe_greeklish(row_data['Phase'])
                    
                    file_name = f"Reports/{latin_name}_{latin_phase}.pdf"
                    
                    zip_file.writestr(file_name, pdf_bytes)
                    progress_bar.progress((index + 1) / len(final_processing_df))
                    row_dict = row.to_dict()
                    
                    evaluations = evaluate_results(row_dict)
                    
                    row_data = {}
                    for k, v in row_dict.items():
                        if pd.isna(v):
                            row_data[k] = "-"
                        else:
                            val_str = str(v).strip()
                            # Αν είναι η ημερομηνία και περιέχει ώρα (π.χ. 2025-09-24 00:00:00), κράτα μόνο την ημερομηνία
                            if k == "Date" and " " in val_str:
                                val_str = val_str.split(" ")[0]
                            row_data[k] = val_str
                    row_data.update(evaluations)

                    # Δημιουργία των <img> tags (εφόσον υπάρχουν οι εικόνες)
                    row_data['Logo1_Tag'] = f'<img src="{LOGO1_PATH}" height="55">' if os.path.exists(LOGO1_PATH) else ''
                    row_data['Logo2_Tag'] = f'<img src="{LOGO2_PATH}" height="55">' if os.path.exists(LOGO2_PATH) else ''
                    
                    for k, v in row_data.items():
                        if k not in evaluations and str(v).replace('.', '', 1).isdigit():
                            try:
                                row_data[k] = f"{float(v):.2f}".rstrip('0').rstrip('.')
                            except ValueError:
                                pass
                    
                    html_content = HTML_TEMPLATE
                    
                    for k, v in row_data.items():
                        html_content = html_content.replace("{" + k + "}", str(v))
                    
                    pdf_io = io.BytesIO()
                    pisa.CreatePDF(html_content, dest=pdf_io, encoding='utf-8', link_callback=link_callback)
                    pdf_bytes = pdf_io.getvalue()
                    
                    latin_name = safe_greeklish(row_data['Name'])
                    latin_phase = safe_greeklish(row_data['Phase'])
                    
                    file_name = f"Reports/{latin_name}_{latin_phase}.pdf"
                    
                    zip_file.writestr(file_name, pdf_bytes)
                    progress_bar.progress((index + 1) / len(final_processing_df))
                    
            st.success("✅ Όλα τα PDF δημιουργήθηκαν με επιτυχία!")
            
            st.download_button(
                label="📥 Λήψη όλων των PDF (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Patient_Reports.zip",
                mime="application/zip"
            )
            
    except Exception as e:
        st.error(f"Παρουσιάστηκε σφάλμα κατά την επεξεργασία: {e}")
