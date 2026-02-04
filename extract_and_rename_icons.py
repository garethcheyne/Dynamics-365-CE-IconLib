"""
Extract and Rename Dynamics 365 Icons
Extracts SVG icons from WOFF font file and renames them with meaningful entity names.
"""
import os
import sys
import shutil
from fontTools.ttLib import TTFont
from fontTools.pens.svgPathPen import SVGPathPen

# Ensure UTF-8 encoding for console output
sys.stdout.reconfigure(encoding='utf-8') if hasattr(sys.stdout, 'reconfigure') else None

# Configuration
FONT_FILE = 'CRMMDL2.9a8be239d5ecd10b97d8f91920f0df73.woff'
OUTPUT_DIR = 'icons'
UNDOCUMENTED_DIR = 'icons/undocumented'

# Entity name to Unicode mappings from styles.css
# Includes both .entity-symbol.Name and .Name-symbol patterns
ENTITY_MAPPINGS = {
    # Entity symbols
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
    # Action symbols from .Name-symbol:before pattern
    'Abandon': 'EA39',
    'Accept': 'E8FB',
    'Activate': 'EFBC',
    'Add': 'EAEE',
    'AddExisting': 'EFF2',
    'AddFriend': 'E1E2',
    'AlignWithFiscalPeriod': 'F6C2',
    'ApplyFilter': 'E8FB',
    'ApplyTemplate': 'E8FF',
    'ApproveArticle': 'F638',
    'ArchiveArticle': 'F635',
    'Arrow': 'E76C',
    'Article': 'F058',
    'ArticleLink': 'F047',
    'Assign': 'F054',
    'AssociateChildCase': 'F072',
    'Attach': 'E723',
    'BackButton': 'E72B',
    'BackButtonWithoutBorder': 'E72B',
    'Cancel': 'E711',
    'CancelCase': 'EFB7',
    'CancelFilter': 'E711',
    'Case': 'E90F',
    'Chat': 'F70F',
    'CheckMark': 'E8FB',
    'CheckedMail': 'ED81',
    'Clear': 'E74D',
    'Close': 'E711',
    'CloseGoal': 'F6C1',
    'ClosePane': 'E89F',
    'Collapsed': 'E76C',
    'Connect': 'EFD4',
    'ContactInfo': 'E779',
    'Convert': 'F05C',
    'ConvertKnowledgeArticle': 'F064',
    'Copy': 'E8C8',
    'CopyLink': 'E8C8',
    'CreateChildCase': 'F06D',
    'CreatePersonalView': 'F554',
    'CSR': 'EDBC',
    'CustomActivity': 'EFF0',
    'CustomEntity': 'EFF7',
    'DeActivate': 'EFB5',
    'Default': 'EB95',
    'Delete': 'E74D',
    'DeleteBulk': 'EFBA',
    'Details': 'E700',
    'Diamond': 'F03D',
    'DiscardArticle': 'F636',
    'Disqualify': 'EFFA',
    'DocumentTemplates': 'EF1C',
    'DownArrow': 'E74B',
    'Drilldown': 'F040',
    'DropdownArrow': 'E70D',
    'Dynamics365': 'EDCC',
    'Edit': 'E70F',
    'EmailIncoming': 'EFAB',
    'EmailLink': 'EFAC',
    'EmailOutgoing': 'EFAF',
    'ExitButton': 'E711',
    'Expanded': 'E70D',
    'ExpandTile': 'F052',
    'ExportToExcel': 'EC28',
    'FailedMail': 'EFAD',
    'FaxIncoming': 'F0D9',
    'FaxOutgoing': 'F0DA',
    'Filter': 'E71C',
    'Find': 'E11A',
    'FinishStage': 'E73E',
    'FirstPageButton': 'F438',
    'FlsLocked': 'E72E',
    'Forward': 'E72A',
    'ForwardButton': 'E761',
    'ForwardEmail': 'E72A',
    'GlobalFilter': 'E71C',
    'GlobalFilterCollapse': '25B7',
    'GlobalFilterExpand': 'EE62',
    'Graph': 'ECB3',
    'HandClick': 'E7C9',
    'Help': 'E897',
    'HighPriority': 'E8C9',
    'Home': 'E80F',
    'Import': 'E150',
    'InsertKbArticle': 'F029',
    'LeftArrowHead': 'E76B',
    'LetterIncoming': 'F0DB',
    'LetterOutgoing': 'F0DC',
    'LinkArticle': 'E71B',
    'Lock': 'E72E',
    'Locked': 'E72E',
    'LowPriority': 'EC4D',
    'Mail': 'E715',
    'Major': 'F066',
    'Manage': 'E912',
    'MarkAsLost': 'E733',
    'MarkAsWon': 'E9D1',
    'MembersIcon': 'E779',
    'MergeRecords': 'F06E',
    'Minor': 'F065',
    'More': 'E712',
    'MoreOptions': 'E9D5',
    'NewAppointment': 'E787',
    'NewEmail': 'EF61',
    'NewFax': 'EF5C',
    'NewLetter': 'ECEF',
    'NewPhoneCall': 'E717',
    'NewRecurringAppointment': 'EF5D',
    'NewTask': 'EFBC',
    'NormalPriority': 'EC4A',
    'OneNote': 'F04F',
    'OpenDelve': 'ECAF',
    'OpenInBrowser': 'E909',
    'OpenInPowerBI': 'EB22',
    'OpenMailbox': 'F008',
    'OpenPane': 'E8A0',
    'OpenPowerBIReport': 'EAE2',
    'Options': 'E8D0',
    'Phone': 'E717',
    'PhoneCallIncoming': 'E77E',
    'PhoneCallOutgoing': 'E77D',
    'Pin': 'E718',
    'Pinned': 'E840',
    'Placeholder': 'EA86',
    'PopOverButton': 'F038',
    'Post': 'E90A',
    'Process': 'F076',
    'PublishKnowledgeArticle': 'F063',
    'Qualify': 'EFF9',
    'QueueIcon': 'F048',
    'QueueItemDetail': 'F052',
    'QueueItemPick': 'F073',
    'QueueItemRelease': 'F06A',
    'QueueItemRemove': 'F06B',
    'QueueItemRoute': 'F069',
    'RatingEmpty': 'E734',
    'RatingFull': 'E735',
    'Reactivate': 'E72C',
    'ReactivateCase': 'EFB6',
    'ReactivateLead': 'E77E',
    'Recalculate': 'E8EF',
    'Refresh': 'E72C',
    'ReduceTile': 'E976',
    'RelateArticle': 'F05E',
    'RelateProduct': 'F060',
    'Remove': 'E74D',
    'RemoveFilter': 'EA49',
    'ReOpenOpportunity': 'E8DE',
    'ReplyAllEmail': 'EE0A',
    'ReplyEmail': 'E97A',
    'Resolve': 'F005',
    'ResolveCase': 'EFB4',
    'RestoreArticle': 'F637',
    'RevertToDraftArticle': 'F644',
    'RunRoutingRule': 'EEEE',
    'SalesLiterature': 'F058',
    'Save': 'E74E',
    'SaveAndClose': 'F038',
    'SaveAndEdit': 'E792',
    'SaveAndRunRoutingRule': 'ED18',
    'SaveAsComplete': 'E8FB',
    'SaveFilterToCurrentPersonalView': 'E74E',
    'SaveFilterToNewPersonalView': 'E792',
    'ScrollLeft': 'E76B',
    'ScrollRight': '00BB',
    'SearchButton': 'E721',
    'SelectChart': 'EFEA',
    'SelectView': 'EC9B',
    'SendDirectEmail': 'E78D',
    'SendEmail': 'E724',
    'SendSelected': 'EFB8',
    'SetActiveButton': 'E7C1',
    'SetAsDefault': 'E739',
    'SetAsHome': 'F00E',
    'SetRegarding': 'EFB9',
    'Settings': 'E713',
    'Share': 'E72D',
    'SiteMap': 'E700',
    'SocialActivityIncoming': 'ECB5',
    'SocialActivityOutgoing': 'ECB5',
    'StageAdvance': 'E72A',
    'StreamView': 'E8BF',
    'SwitchProcess': 'E8B1',
    'SystemPost': 'EFEF',
    'TileView': 'E8A9',
    'Tools': 'ED15',
    'Translate': 'F2E6',
    'UnlinkArticle': 'ED90',
    'Unpin': 'E77A',
    'UpdateArticle': 'F068',
    'UpArrow': 'E74A',
    'UpArrowHead': 'E70E',
    'ViewNotifications': 'E7C1',
    'VisualFilter': 'E9EC',
    'WordTemplates': 'EEF3',
}


def extract_icons_from_font():
    """Extract all glyphs from WOFF font as SVG files."""
    print(f"Loading font: {FONT_FILE}")
    font = TTFont(FONT_FILE)
    
    # Clean and recreate output directory to avoid leftover icons
    if os.path.exists(OUTPUT_DIR):
        print(f"Cleaning existing '{OUTPUT_DIR}/' folder...")
        shutil.rmtree(OUTPUT_DIR)
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
    
    # Create subdirectory for undocumented icons
    os.makedirs(UNDOCUMENTED_DIR, exist_ok=True)
    
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
            
            # Save with entity name(s) to main icons directory
            for entity_name in entity_names:
                filename = f"{entity_name}.svg"
                filepath = os.path.join(OUTPUT_DIR, filename)
                
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(svg_content)
                
                all_files.append(filename)
                saved_count += 1
                print(f"✓ {filename}")
        else:
            # Save with unicode name to undocumented subdirectory
            filename = f"icon_U+{unicode_hex}.svg"
            filepath = os.path.join(UNDOCUMENTED_DIR, filename)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            all_files.append(f"undocumented/{filename}")
            unnamed_count += 1
    
    print(f"\n{'='*60}")
    print(f"✓ Saved {saved_count} named icons to '{OUTPUT_DIR}/'")
    print(f"✓ Saved {unnamed_count} undocumented icons to '{UNDOCUMENTED_DIR}/'")
    print(f"📁 Total: {saved_count + unnamed_count} icons")
    print(f"{'='*60}\n")
    
    if unnamed_count > 0:
        print(f"Note: {unnamed_count} icons are not defined in CSS (undocumented).")
        print("These are stored separately in the 'undocumented' folder.")
    
    return all_files


def create_readme(icon_files):
    """Create a README.md file with an icon matrix."""
    print("\nGenerating README.md with icon matrix...")
    
    # Sort icons alphabetically (separate named and undocumented)
    named_icons = sorted([f for f in icon_files if not f.startswith('undocumented/')])
    undocumented_icons = sorted([f for f in icon_files if f.startswith('undocumented/')])
    
    # Create README content
    readme = """# Dynamics 365 Icon Library

This library contains **{total}** icons extracted from the Dynamics 365 CRM MDL2 font.

- **{named} documented icons** with semantic names (Account, Contact, OpenInBrowser, etc.)
- **{undoc} undocumented icons** not defined in CSS (stored in `undocumented/` folder)

## Named Entity Icons ({named} icons)

Icons with semantic entity names from Dynamics 365 CSS definitions.

| Icon | Name | Preview |
|------|------|---------|
""".format(total=len(icon_files), named=len(named_icons), undoc=len(undocumented_icons))
    
    # Add named icons
    for icon_file in named_icons:
        name = icon_file.replace('.svg', '')
        readme += f"| ![{name}](icons/{icon_file}) | **{name}** | [View](icons/{icon_file}) |\n"
    
    # Add Undocumented icons section
    readme += f"""
## Undocumented Icons ({len(undocumented_icons)} icons)

Additional icons in the font that are **not defined in CSS**. These may be internal, deprecated, or future icons.
Stored in the `icons/undocumented/` folder.

<details>
<summary>Click to expand undocumented icon list (first 100)</summary>

| Icon | Unicode | Preview |
|------|---------|---------|
"""
    
    for icon_file in undocumented_icons[:100]:  # Show first 100 in README
        name = icon_file.replace('undocumented/', '').replace('.svg', '').replace('icon_', '')
        readme += f"| ![{name}](icons/{icon_file}) | **{name}** | [View](icons/{icon_file}) |\n"
    
    if len(undocumented_icons) > 100:
        readme += f"\n*... and {len(undocumented_icons) - 100} more undocumented icons*\n"
    
    readme += """
</details>

## Usage

### In HTML
```html
<img src="icons/Account.svg" alt="Account" width="24" height="24">
<!-- Or undocumented icons -->
<img src="icons/undocumented/icon_U+E700.svg" alt="Icon" width="24" height="24">
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
    
    print(f"✓ Created README.md with {len(icon_files)} icons")


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
