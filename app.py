#!/usr/bin/env python3
"""
QX Matrix Interactive Analyzer - Optimized Streamlit version
Optimizations: Caching, lazy loading, vectorized operations
"""

import pandas as pd
import streamlit as st
from pathlib import Path
from io import BytesIO
import time

# Page configuration
st.set_page_config(
    page_title="QX Matrix Analyzer",
    page_icon="🔍",
    layout="wide"
)


class QXMatrixAnalyzer:
    """Optimized analyzer for QX matrix Excel files"""

    def __init__(self, excel_file):
        """Initialize analyzer with an Excel file"""
        self.excel_file = excel_file
        self.data = None
        self._xlrd_book = None
        self._xlrd_sheet = None
        
    @st.cache_data(show_spinner=False)
    def load_data(_self, file_bytes):
        """Load Excel file - cached for performance"""
        try:
            # Try openpyxl first (for .xlsx)
            data = pd.read_excel(BytesIO(file_bytes), header=None, engine='openpyxl')
            engine = 'openpyxl'
        except Exception:
            try:
                # Try xlrd for old .xls files
                data = pd.read_excel(BytesIO(file_bytes), header=None, engine='xlrd')
                engine = 'xlrd'
            except Exception as e:
                st.error(f"Unable to load file: {e}")
                return None, None
        return data, engine
    
    def _get_xlrd_objects(self):
        """Lazy load xlrd objects only when needed for color information"""
        if self._xlrd_book is None:
            try:
                import xlrd
                self._xlrd_book = xlrd.open_workbook(
                    file_contents=self.excel_file.read(),
                    formatting_info=True
                )
                self._xlrd_sheet = self._xlrd_book.sheet_by_index(0)
                self.excel_file.seek(0)  # Reset file pointer
            except Exception as e:
                st.warning(f"Color information not available: {e}")
                return None, None
        return self._xlrd_book, self._xlrd_sheet
    
    def get_cell_color(self, row_name, col_idx):
        """Get cell background color"""
        book, sheet = self._get_xlrd_objects()
        if not book or not sheet:
            return None, None
            
        try:
            import xlrd
            # Find row index for the subset
            row_idx = None
            for r in range(3, 15):
                cell = sheet.cell(r, 75)
                if cell.ctype == xlrd.XL_CELL_TEXT and cell.value.strip() == row_name:
                    row_idx = r
                    break
            
            if row_idx is None:
                return None, None
                
            xf_index = sheet.cell_xf_index(row_idx, col_idx)
            xf = book.xf_list[xf_index]
            bg_color_index = xf.background.pattern_colour_index
            palette = book.colour_map
            rgb = palette.get(bg_color_index)
            return bg_color_index, rgb
        except Exception as e:
            return None, f"Error: {e}"
    
    def get_component_color(self, row_name, col_idx):
        """Get component cell background color"""
        return self.get_cell_color(row_name, col_idx)
    
    def find_defect_column(self, defect_name):
        """Find column index for a specific defect type - optimized"""
        if self.data is None:
            return None
            
        defect_row_idx = 15
        if defect_row_idx >= len(self.data):
            return None
        
        # Vectorized search
        defect_row = self.data.iloc[defect_row_idx]
        mask = defect_row.notna() & defect_row.astype(str).str.lower().str.contains(defect_name.lower(), na=False)
        matches = defect_row[mask]
        
        return matches.index[0] if len(matches) > 0 else None
    
    def list_all_defects(self):
        """List all defect types - optimized with vectorization"""
        if self.data is None:
            return []
            
        defect_row_idx = 15
        if defect_row_idx >= len(self.data):
            return []
        
        defect_row = self.data.iloc[defect_row_idx]
        
        # Vectorized filtering
        valid_mask = defect_row.notna() & (defect_row.astype(str).str.strip() != '')
        valid_defects = defect_row[valid_mask]
        
        # Filter to start from column 78 (CA in Excel)
        valid_defects = valid_defects[valid_defects.index >= 78]
        
        return [(col, str(val).strip()) for col, val in valid_defects.items()]
    
    def get_sous_ensembles(self, defect_col, defect_name=None):
        """Get all subsets associated with a defect column"""
        sous_ensembles = []
        

        
        # Rule: each line 3-14 where defect column equals 'ç' is a subset
        target_marker = 'ç'
        for row_idx in range(3, 15):
            if row_idx < len(self.data) and defect_col < len(self.data.columns):
                cell = self.data.iloc[row_idx, defect_col]
                if pd.notna(cell) and str(cell).strip() == target_marker:
                    sub_name = self.data.iloc[row_idx, 75] if 75 < len(self.data.columns) else None
                    if pd.notna(sub_name):
                        name = str(sub_name).strip()
                        if name and name not in sous_ensembles:
                            sous_ensembles.append(name)
        
        # Fallback: if nothing found, try (13, 75)
        if not sous_ensembles:
            sub_name = self.data.iloc[13, 75] if 75 < len(self.data.columns) else None
            if pd.notna(sub_name):
                sous_ensembles.append(str(sub_name).strip())
        
        return sous_ensembles
    
    def get_composants(self, sous_ensemble_name):
        """Get all components belonging to a sub-assembly"""
        components = []
        
        # Find the row with this sub-assembly name
        sub_row_idx = None
        for row_idx in range(3, 14):
            sub_name = self.data.iloc[row_idx, 75] if 75 < len(self.data.columns) else None
            if pd.notna(sub_name) and str(sub_name).strip() == sous_ensemble_name:
                sub_row_idx = row_idx
                break
        
        if sub_row_idx is None:
            return components
        
        # Find columns with markers in this row
        markers = ['ê', 'è', 'ç', 'ª']
        for col in range(0, 74):  # Components are in columns 0-73
            if col < len(self.data.columns):
                marker = str(self.data.iloc[sub_row_idx, col]).strip()
                if marker in markers:
                    # Get component name from row 15
                    component = self.data.iloc[15, col] if col < len(self.data.columns) else None
                    if pd.notna(component):
                        comp_name = str(component).strip()
                        if comp_name not in markers and comp_name != '':
                            components.append(comp_name)
        
        return components
    
    def get_parametres(self, sous_ensemble_name):
        """
        Get parameters linked to components in a sub-assembly.
        
        Args:
            sous_ensemble_name: Name of the sub-assembly
        
        Returns:
            List of dictionaries with parameter info and linked components
        """
        parameters = []
        
        if not sous_ensemble_name:
            return parameters
        
        # First, find all component columns for this sub-assembly
        component_cols = []
        sub_row_idx = None
        
        # Find the row with this sub-assembly
        for row_idx in range(3, 14):
            sub_name = self.data.iloc[row_idx, 75] if 75 < len(self.data.columns) else None
            if pd.notna(sub_name) and str(sub_name).strip() == sous_ensemble_name:
                sub_row_idx = row_idx
                break
        
        if sub_row_idx is None:
            return parameters
        
        # Find component columns with markers in this sub-assembly row
        markers = ['ê', 'è', 'ç', 'ª']
        for col in range(0, 74):
            if col < len(self.data.columns):
                marker = str(self.data.iloc[sub_row_idx, col]).strip()
                if marker in markers:
                    component_cols.append(col)
        
        # Now find parameters linked to these components
        for row_idx in range(17, len(self.data)):
            param_name = self.data.iloc[row_idx, 75] if 75 < len(self.data.columns) else None
            param_value = self.data.iloc[row_idx, 76] if 76 < len(self.data.columns) else None
            
            if pd.notna(param_name):
                param_name_str = str(param_name).strip()
                param_value_str = str(param_value).strip() if pd.notna(param_value) else ""
                
                # Check if any component in our sub-assembly has a marker in this row
                linked_components = []
                for col in component_cols:
                    if col < len(self.data.columns):
                        cell = self.data.iloc[row_idx, col]
                        if pd.notna(cell) and str(cell).strip() in ['è', 'ê']:
                            # Get component name from row 15
                            comp_name = self.data.iloc[15, col] if col < len(self.data.columns) else None
                            if pd.notna(comp_name):
                                comp_name_str = str(comp_name).strip()
                                if comp_name_str not in markers and comp_name_str != '':
                                    linked_components.append(comp_name_str)
                
                if linked_components and param_name_str not in markers and param_name_str != '':
                    parameters.append({
                        'name': param_name_str,
                        'value': param_value_str,
                        'components': linked_components
                    })
        
        return parameters
    
    def analyze_defect(self, defect_name):
        """Analyse complète d'un type de défaut."""
        # Find the defect column
        defect_col = self.find_defect_column(defect_name)
        
        if defect_col is None:
            return {
                'error': f'Défaut "{defect_name}" introuvable dans la matrice',
                'defect_name': defect_name
            }
        
        # Get sous-ensembles (avec le nom du défaut pour corrections manuelles)
        sous_ensembles = self.get_sous_ensembles(defect_col, defect_name)
        
        # Get composants et paramètres pour chaque sous-ensemble
        composants = []
        parametres = []
        for se in sous_ensembles:
            composants.extend(self.get_composants(se))
            parametres.extend(self.get_parametres(se))
        
        # Déduplication des composants en conservant l'ordre
        seen = set()
        composants_uniques = []
        for comp in composants:
            if comp not in seen:
                seen.add(comp)
                composants_uniques.append(comp)
        
        return {
            'defect_name': defect_name,
            'defect_column': defect_col,
            'sous_ensembles': sous_ensembles,
            'composants': composants_uniques,
            'parametres': parametres
        }
    
    def get_filtered_hierarchy(self, defect_name):
        """Retourne la hiérarchie filtrée par couleur pour un défaut donné."""
        result = self.analyze_defect(defect_name)
        
        if 'error' in result:
            return {
                'defect_name': defect_name,
                'defect_column': None,
                'sous_ensembles': [],
                'bonnes_composants': [],
                'parametres': []
            }
        
        # Section 2 - Sous-ensembles
        sous_ensembles = result['sous_ensembles']
        defect_col = result['defect_column']
        
        # Section 3 - Bonnes composants
        bonnes_composants = []
        color_cache = {}
        
        if sous_ensembles:
            for se in sous_ensembles:
                if se not in color_cache:
                    color_idx, _ = self.get_cell_color(se, defect_col)
                    color_cache[se] = color_idx
                
                color_idx = color_cache[se]
                
                for comp in result['composants']:
                    comp_cols = []
                    comp_name = comp.strip()
                    
                    for col in range(0, 74):
                        cell_val = self.data.iloc[15, col]
                        if pd.notna(cell_val) and str(cell_val).strip() == comp_name:
                            comp_cols.append(col)
                    
                    for comp_col in comp_cols:
                        comp_color_idx, _ = self.get_component_color(se, comp_col)
                        if comp_color_idx == color_idx:
                            bonnes_composants.append(comp)
        
        # Déduplication
        bonnes_composants_uniques = []
        seen = set()
        for comp in bonnes_composants:
            if comp not in seen:
                seen.add(comp)
                bonnes_composants_uniques.append(comp)
        
        # Section 4 - Paramètres pour les bonnes composants
        parametres_filtres = []
        bons_comp_set = set(bonnes_composants_uniques)
        
        for param in result['parametres']:
            linked_in_good = [c for c in param['components'] if c in bons_comp_set]
            if linked_in_good:
                parametres_filtres.append({
                    'name': param['name'],
                    'value': param['value'],
                    'components': linked_in_good
                })
        
        return {
            'defect_name': defect_name,
            'defect_column': defect_col,
            'sous_ensembles': sous_ensembles,
            'bonnes_composants': bonnes_composants_uniques,
            'parametres': parametres_filtres
        }


def display_hierarchy(result):
    """Display hierarchy in Streamlit with nice formatting"""
    st.markdown("---")
    
    # Type de défaut
    st.markdown("### 🎯 Type de défaut")
    col1, col2 = st.columns([3, 1])
    with col1:
        st.info(f"**{result['defect_name']}**")
    with col2:
        if result['defect_column'] is not None:
            st.metric("Colonne", result['defect_column'])
    
    # Sous-ensembles
    st.markdown("### 📦 Sous-ensembles")
    if result['sous_ensembles']:
        for se in result['sous_ensembles']:
            st.write(f"• {se}")
    else:
        st.warning("Aucun sous-ensemble trouvé")
    
    # Composants
    st.markdown(f"### ⚙️ Composants ({len(result['bonnes_composants'])})")
    if result['bonnes_composants']:
        # Display in columns for better layout
        num_cols = 3
        cols = st.columns(num_cols)
        for idx, comp in enumerate(result['bonnes_composants']):
            with cols[idx % num_cols]:
                st.write(f"{idx + 1}. {comp}")
    else:
        st.warning("Aucun composant trouvé")
    
    # Paramètres
    st.markdown(f"### 🔧 Paramètres ({len(result['parametres'])})")
    if result['parametres']:
        for param in result['parametres']:
            with st.expander(f"**{param['name']}**", expanded=False):
                if param['value']:
                    st.write(f"**Action:** {param['value']}")
                st.write(f"**Composants liés ({len(param['components'])}):**")
                st.write(", ".join(param['components']))
    else:
        st.warning("Aucun paramètre trouvé")


def export_to_pdf(result):
    """Export hierarchy to PDF format"""
    try:
        from fpdf import FPDF
        
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Title
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, f"Hierarchie pour: {result['defect_name']}", ln=True, align='C')
        pdf.ln(5)
        
        # Section 1 - Type de défaut
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, "1 - Type de defaut", ln=True)
        pdf.set_font("Arial", '', 12)
        pdf.cell(0, 8, f"Nom: {result['defect_name']}", ln=True)
        if result['defect_column'] is not None:
            pdf.cell(0, 8, f"Colonne: {result['defect_column']}", ln=True)
        pdf.ln(5)
        
        # Section 2 - Sous-ensembles
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, "2 - Sous-ensembles", ln=True)
        pdf.set_font("Arial", '', 12)
        if result['sous_ensembles']:
            for se in result['sous_ensembles']:
                pdf.cell(0, 8, f"  - {se}", ln=True)
        else:
            pdf.cell(0, 8, "  Non trouve", ln=True)
        pdf.ln(5)
        
        # Section 3 - Composants
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"3 - Bonnes composants ({len(result['bonnes_composants'])})", ln=True)
        pdf.set_font("Arial", '', 12)
        if result['bonnes_composants']:
            for i, comp in enumerate(result['bonnes_composants'], 1):
                # Handle long component names
                pdf.multi_cell(0, 8, f"  {i}. {comp}")
        else:
            pdf.cell(0, 8, "  (Aucun composant trouve)", ln=True)
        pdf.ln(5)
        
        # Section 4 - Paramètres
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"4 - Parametres composant ({len(result['parametres'])})", ln=True)
        pdf.set_font("Arial", '', 12)
        if result['parametres']:
            for param in result['parametres']:
                pdf.set_font("Arial", 'B', 12)
                pdf.multi_cell(0, 8, f"  * {param['name']}")
                pdf.set_font("Arial", '', 12)
                if param['value']:
                    pdf.multi_cell(0, 8, f"    Action: {param['value']}")
                comp_str = ', '.join(param['components'])
                pdf.multi_cell(0, 8, f"    Composants lies: {comp_str}")
                pdf.ln(3)
        else:
            pdf.cell(0, 8, "  (Aucun parametre trouve)", ln=True)
        
        # Footer
        pdf.ln(10)
        pdf.set_font("Arial", 'I', 10)
        pdf.cell(0, 8, f"Genere automatiquement depuis: {result['defect_name']}", ln=True, align='C')
        
        # Return PDF as bytes
        pdf_output = pdf.output(dest='S')
        if isinstance(pdf_output, str):
            return pdf_output.encode('latin-1')
        return pdf_output
    except ImportError:
        return None


def main():
    """Main Streamlit application"""
    st.title("🔍 QX Matrix Analyzer")
    
    # Sidebar
    with st.sidebar:
        st.header("📂 Chargement du fichier")
        uploaded_file = st.file_uploader(
            "Choisir un fichier Excel (.xls ou .xlsx)",
            type=['xls', 'xlsx']
        )
        
        if uploaded_file:
            st.success(f"✓ Fichier chargé: {uploaded_file.name}")
            st.info(f"Taille: {uploaded_file.size / 1024:.1f} KB")
    
    # Main content
    if uploaded_file is None:
        st.info("👈 Veuillez charger un fichier Excel pour commencer")
        st.markdown("""
        ### À propos
        Cet outil permet d'analyser les matrices QX et d'extraire:
        - Types de défauts
        - Sous-ensembles associés
        - Composants concernés
        - Paramètres liés
        
        ### Optimisations
        - ⚡ Mise en cache des données
        - 🚀 Opérations vectorisées
        - 💾 Chargement paresseux des couleurs
        """)
        return
    
    # Initialize analyzer
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)  # Reset file pointer
    
    with st.spinner("Chargement des données..."):
        analyzer = QXMatrixAnalyzer(uploaded_file)
        data, engine = analyzer.load_data(file_bytes)
        
        if data is None:
            st.error("Impossible de charger le fichier")
            return
        
        analyzer.data = data
        st.success(f"✓ Chargé avec {engine} | {len(data)} lignes × {len(data.columns)} colonnes")
    
    # Tabs for different functionalities
    tab1, tab2, tab3 = st.tabs(["📋 Liste des défauts", "🔍 Analyser un défaut", "🔎 Recherche"])
    
    with tab1:
        st.header("Liste de tous les types de défaut")
        
        with st.spinner("Chargement des défauts..."):
            start = time.time()
            defects = analyzer.list_all_defects()
            elapsed = time.time() - start
        
        st.metric("Nombre de défauts", len(defects))
        
        if defects:
            # Create a DataFrame for better display
            df = pd.DataFrame([d[1] for d in defects], columns=['Type de défaut'])
            
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )
            
            # Download button
            csv = df.to_csv(index=False)
            st.download_button(
                label="📥 Télécharger en CSV",
                data=csv,
                file_name="defects_list.csv",
                mime="text/csv"
            )
    
    with tab2:
        st.header("Analyser un défaut spécifique")
        
        # Get list of defects for autocomplete
        defects = analyzer.list_all_defects()
        defect_names = [d[1] for d in defects]
        
        selected_defect = st.selectbox(
            "Sélectionner un type de défaut",
            options=[""] + defect_names,
            index=0
        )
        
        if selected_defect:
            with st.spinner("Analyse en cours..."):
                start = time.time()
                result = analyzer.get_filtered_hierarchy(selected_defect)
                elapsed = time.time() - start
            
            st.success(f"✓ Analyse terminée en {elapsed*1000:.1f} ms")
            
            display_hierarchy(result)
    
    with tab3:
        st.header("Recherche par mot-clé")
        
        keyword = st.text_input("🔎 Entrer un mot-clé", "")
        
        if keyword:
            defects = analyzer.list_all_defects()
            matches = [(col, d) for col, d in defects if keyword.lower() in d.lower()]
            
            if matches:
                st.success(f"✓ {len(matches)} défaut(s) trouvé(s)")
                
                for col, defect in matches:
                    if st.button(f"🎯 {defect}", key=f"match_{col}"):
                        with st.spinner("Analyse en cours..."):
                            result = analyzer.get_filtered_hierarchy(defect)
                        display_hierarchy(result)
            else:
                st.warning(f"Aucun défaut trouvé contenant '{keyword}'")


if __name__ == "__main__":
    main()