"""
Extract and Rename Dynamics 365 Icons
Extracts SVG icons from WOFF font file and renames them with meaningful entity names.
"""
import os
import shutil
from fontTools.ttLib import TTFont
from fontTools.pens.svgPathPen import SVGPathPen

# Configuration
FONT_FILE = 'CRMMDL2.9a8be239d5ecd10b97d8f91920f0df73.woff'
OUTPUT_DIR = 'icons'

# Entity name to Unicode mappings from styles.css
ENTITY_MAPPINGS = {
    'Entity': 'ECDC',
    'List': 'EA37',
    'Account': 'EED6',
    'Opportunity': 'F05F',
    'Timer': 'E7C1',
    'Sharepointdocument': 'E7C3',
    'Dashboard': 'E9FE',
    'WORKSPACE': 'E9FE',
    'Lead': 'EFD6',
    'Contact': 'E8D4',
    'Activitypointer': 'EFF4',
    'Drafts': 'F05B',
    'Systemuser': 'E77B',
    'Letter': 'ECEF',
    'Salesorder': 'F03B',
    'Calendar': 'E787',
    'Category': 'ECA6',
    'Competitor': 'EE57',
    'Task': 'EADF',
    'Fax': 'EF5C',
    'Email': 'E715',
    'Entitlement': 'EB95',
    'Phonecall': 'E717',
    'Contract': 'EB95',
    'Quote': 'F067',
    'Incident': 'E7B9',
    'Campaign': 'E789',
    'Connection': 'EFD4',
    'CustomerAddress': 'EEBD',
    'Position': 'ECA6',
    'TransactionCurrency': 'EAE4',
    'Appointment': 'E787',
    'Team': 'E902',
    'Invoice': 'EFE4',
    'Knowledgearticle': 'F000',
    'Product': 'ECDC',
    'Opportunityproduct': 'EFE8',
    'Queue': 'EFBF',
    'Queueitem': 'EFBF',
    'Socialprofile': 'ECFE',
    'ChevronRight': 'F06C',
    'Globe': 'E774',
    'Ticker': 'E7BB',
    'Duration': 'E91E',
    'Timezone': 'EBF9',
    'Language': 'F039',
    'MultipleUsers': 'E716',
    'Regarding': 'E71B',
    'Checklist': 'F072',
    'TwoOptions': 'E7B6',
    'Currency': 'EB0D',
    'DateTime': 'EC92',
    'OfficeIcon': 'EC29',
    'Service': 'EFD2',
    'ServiceAppointment': 'EFF1',
    'Equipment': 'F426',
    'BusinessUnit': 'E821',
    'PriceLevel': 'EFC6',
    'UoMSchedule': 'EDEC',
    'DiscountType': 'EB07',
    'Territory': 'E800',
    'Socialactivity': 'E8F2',
    'Code': 'E943',
    'CompletedSolid': 'EC61',
    'WarningSolid': 'F736',
    'Site': 'E731',
}


def extract_icons_from_font():
    """Extract all glyphs from WOFF font as SVG files."""
    print(f"Loading font: {FONT_FILE}")
    font = TTFont(FONT_FILE)
    
    # Create output directory
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    glyf_table = font['glyf']
    cmap = font.getBestCmap()
    
    print(f"Found {len(cmap)} glyphs in font")
    print(f"Extracting to '{OUTPUT_DIR}/' folder...\n")
    
    extracted = {}
    
    for unicode_value, glyph_name in cmap.items():
        try:
            pen = SVGPathPen(font.getGlyphSet())
            font.getGlyphSet()[glyph_name].draw(pen)
            path_data = pen.getCommands()
            
            if path_data:
                # Get glyph metrics for viewBox
                glyph = glyf_table[glyph_name]
                if hasattr(glyph, 'xMin'):
                    view_box = f"{glyph.xMin} {-glyph.yMax} {glyph.xMax - glyph.xMin} {glyph.yMax - glyph.yMin}"
                else:
                    view_box = "0 0 1000 1000"
                
                # Create SVG content
                svg_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="{view_box}">
  <path d="{path_data}" transform="scale(1,-1)"/>
</svg>'''
                
                # Store by unicode hex value
                unicode_hex = f"{unicode_value:04X}".upper()
                extracted[unicode_hex] = svg_content
        
        except Exception as e:
            pass  # Skip glyphs that can't be extracted
    
    print(f"✓ Extracted {len(extracted)} SVG icons\n")
    return extracted


def save_icons_with_names(extracted_icons):
    """Save icons with meaningful entity names."""
    print("Saving icons with entity names...\n")
    
    saved_count = 0
    unnamed_count = 0
    all_files = []
    
    # Create reverse mapping: unicode -> entity_names
    unicode_to_entities = {}
    for entity_name, unicode_hex in ENTITY_MAPPINGS.items():
        unicode_upper = unicode_hex.upper()
        if unicode_upper not in unicode_to_entities:
            unicode_to_entities[unicode_upper] = []
        unicode_to_entities[unicode_upper].append(entity_name)
    
    # Save named icons
    for unicode_hex, svg_content in extracted_icons.items():
        if unicode_hex in unicode_to_entities:
            entity_names = unicode_to_entities[unicode_hex]
            
            # Save with entity name(s)
            for entity_name in entity_names:
                filename = f"{entity_name}.svg"
                filepath = os.path.join(OUTPUT_DIR, filename)
                
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(svg_content)
                
                all_files.append(filename)
                saved_count += 1
                print(f"✓ {filename}")
        else:
            # Save with unicode name for unmapped icons
            filename = f"icon_U+{unicode_hex}.svg"
            filepath = os.path.join(OUTPUT_DIR, filename)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            all_files.append(filename)
            unnamed_count += 1
    
    print(f"\n{'='*60}")
    print(f"✓ Saved {saved_count} named icons")
    print(f"✓ Saved {unnamed_count} icons with Unicode names")
    print(f"📁 Total: {saved_count + unnamed_count} icons in '{OUTPUT_DIR}/'")
    print(f"{'='*60}\n")
    
    if unnamed_count > 0:
        print(f"Note: {unnamed_count} icons don't have entity name mappings.")
        print("You can add more mappings to ENTITY_MAPPINGS in this script.")
    
    return all_files


def create_readme(icon_files):
    """Create a README.md file with an icon matrix."""
    print("\nGenerating README.md with icon matrix...")
    
    # Sort icons alphabetically
    sorted_icons = sorted(icon_files)
    
    # Create README content
    readme = """# Dynamics 365 Icon Library

This library contains **{total}** icons extracted from the Dynamics 365 CRM MDL2 font.

## Named Entity Icons ({named} icons)

Icons with semantic entity names from Dynamics 365.

| Icon | Name | Preview |
|------|------|---------|
""".format(total=len(sorted_icons), named=len([f for f in sorted_icons if not f.startswith('icon_U+')]))
    
    # Add named icons
    named_icons = [f for f in sorted_icons if not f.startswith('icon_U+')]
    for icon_file in named_icons:
        name = icon_file.replace('.svg', '')
        readme += f"| ![{name}](icons/{icon_file}) | **{name}** | [View](icons/{icon_file}) |\n"
    
    # Add Unicode icons section
    readme += f"""
## Unicode Icons ({len(sorted_icons) - len(named_icons)} icons)

Additional icons identified by Unicode value.

<details>
<summary>Click to expand full Unicode icon list</summary>

| Icon | Unicode | Preview |
|------|---------|---------|
"""
    
    unicode_icons = [f for f in sorted_icons if f.startswith('icon_U+')]
    for icon_file in unicode_icons[:100]:  # Show first 100 in README
        name = icon_file.replace('.svg', '').replace('icon_', '')
        readme += f"| ![{name}](icons/{icon_file}) | **{name}** | [View](icons/{icon_file}) |\n"
    
    if len(unicode_icons) > 100:
        readme += f"\n*... and {len(unicode_icons) - 100} more Unicode icons*\n"
    
    readme += """
</details>

## Usage

### In HTML
```html
<img src="icons/Account.svg" alt="Account" width="24" height="24">
```

### In CSS
```css
.icon {
  background-image: url('icons/Account.svg');
  width: 24px;
  height: 24px;
}
```

### Direct SVG
```html
<svg width="24" height="24">
  <use href="icons/Account.svg#icon"/>
</svg>
```

## Extraction Process

Icons were extracted from `CRMMDL2.woff` font file using FontTools and mapped to entity names from Dynamics 365 CSS classes.

To re-extract icons:
```bash
pip install fonttools
python extract_and_rename_icons.py
```

## Icon Categories

- **CRM Entities**: Account, Contact, Lead, Opportunity, etc.
- **Activities**: Email, Phone, Task, Appointment, etc.
- **System**: Dashboard, Settings, Calendar, etc.
- **Documents**: Various Office file type icons
- **UI Controls**: Arrows, buttons, navigation elements

## License

These icons are part of Microsoft Dynamics 365 and subject to Microsoft's licensing terms.
"""
    
    # Write README
    with open('README.md', 'w', encoding='utf-8') as f:
        f.write(readme)
    
    print(f"✓ Created README.md with {len(sorted_icons)} icons")


def cleanup_temp_files():
    """Remove temporary files and directories."""
    temp_dirs = ['extracted_icons', 'renamed_icons', 'venv']
    temp_files = [
        'extract_icons.py', 'rename_icons.py', 'rename_svg.py',
        'test_regex.py', 'debug_css.py', 'rename_final.py',
        'rename_entities.py', 'entities.css'
    ]
    
    for dirname in temp_dirs:
        if os.path.exists(dirname):
            try:
                shutil.rmtree(dirname)
                print(f"✓ Removed {dirname}/")
            except:
                pass
    
    for filename in temp_files:
        if os.path.exists(filename):
            try:
                os.remove(filename)
                print(f"✓ Removed {filename}")
            except:
                pass


def main():
    """Main execution function."""
    print("="*60)
    print("Dynamics 365 Icon Extractor")
    print("="*60)
    print()
    
    # Step 1: Extract icons from font
    extracted_icons = extract_icons_from_font()
    
    # Step 2: Save with entity names
    icon_files = save_icons_with_names(extracted_icons)
    
    # Step 3: Create README with icon matrix
    create_readme(icon_files)
    
    # Step 4: Clean up
    print("\nCleaning up temporary files...")
    cleanup_temp_files()
    
    print("\n✓ Done! All icons extracted to 'icons/' folder")
    print("✓ View README.md for icon gallery")


if __name__ == "__main__":
    main()
