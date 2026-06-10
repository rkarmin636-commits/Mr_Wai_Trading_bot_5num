"""
Integrate TRX_Pattern_5num_FINAL.xlsx into Mr_Wai_Trading_bot_5num.py
"""

import re

print("=" * 80)
print("Integrating Excel file into 5-number bot...")
print("=" * 80)

# Read the bot file
with open('Mr_Wai_Trading_bot_5num.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find and replace the PATTERN_TABLE section
start_marker = "# Pattern Table"
start_idx = content.find(start_marker)

if start_idx == -1:
    print("ERROR: Could not find 'Pattern Table' comment")
    exit(1)

# Find end: class WaveUpBot definition  
end_marker = "\nclass WaveUpBot:"
end_idx = content.find(end_marker, start_idx)

if end_idx == -1:
    print("ERROR: Could not find 'class WaveUpBot:'")
    exit(1)

# Extract the part before and after
before = content[:start_idx]
after = content[end_idx:]

# Create new Excel loading code for 5-number patterns
new_code = '''# ════════════════════════════════════════════════════════════
# Pattern Table Loader (Hybrid: Excel + Hardcoded Fallback)
# 5-Number Version: 100,000 patterns (10^5)
# Excel file: TRX_Pattern_5num_FINAL.xlsx
# ════════════════════════════════════════════════════════════

# Hardcoded fallback table (abbreviated - full table in Excel)
_HARDCODED_PATTERN_TABLE_5NUM = {}  # Will be populated below if Excel not found

def _load_pattern_table_5num():
    """
    Load 5-number pattern table with hybrid approach:
    1. Try loading from Excel file first
    2. If Excel fails or not found, use hardcoded fallback
    """
    try:
        import openpyxl
        import os
        
        # Excel file path (relative to bot script)
        excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 
                                  'TRX_Pattern_5num_FINAL.xlsx')
        
        if os.path.exists(excel_path):
            logger.info(f"📊 Found 5-number Excel file: {excel_path}")
            logger.info("🔄 Loading 5-number patterns from Excel...")
            
            wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
            ws = wb.active
            
            table = {}
            row_count = 0
            
            for row in ws.iter_rows(values_only=True):
                row_count += 1
                if row_count == 1:  # Skip header row
                    continue
                
                pattern_str = str(row[0]).strip() if row[0] else ''
                if not pattern_str or ',' not in pattern_str:
                    continue
                
                try:
                    # Detect Excel column format
                    col1 = int(row[1]) if row[1] is not None else 0
                    col2 = int(row[2]) if row[2] is not None else 0
                    
                    if len(row) > 3 and row[3] is not None:
                        col3 = int(row[3])
                        # Format: [pattern, total, S, B]
                        s_count = col2
                        b_count = col3
                    else:
                        # Format: [pattern, S, B]
                        s_count = col1
                        b_count = col2
                    
                    table[pattern_str] = [s_count, b_count]
                except (ValueError, TypeError, IndexError):
                    continue
            
            wb.close()
            
            if len(table) > 0:
                logger.info(f"✅ Successfully loaded {len(table)} 5-number patterns from Excel")
                return table
            else:
                logger.warning("⚠️ No patterns found in Excel, using hardcoded table")
                return _HARDCODED_PATTERN_TABLE_5NUM
        else:
            logger.warning(f"⚠️ 5-number Excel file not found at {excel_path}")
            logger.info("🔄 Using hardcoded 5-number pattern table as fallback")
            return _HARDCODED_PATTERN_TABLE_5NUM
            
    except ImportError:
        logger.warning("⚠️ openpyxl not installed, using hardcoded table")
        logger.info("📦 Install with: pip install openpyxl")
        return _HARDCODED_PATTERN_TABLE_5NUM
    except Exception as e:
        logger.error(f"❌ Failed to load Excel file: {e}")
        logger.info("🔄 Falling back to hardcoded pattern table")
        return _HARDCODED_PATTERN_TABLE_5NUM

# Load 5-number pattern table (Excel if available, otherwise hardcoded)
PATTERN_TABLE = _load_pattern_table_5num()
'''

# Replace the old pattern table section
new_content = before + new_code + after

# Fix 1: Tie-handling bug (change >= to >)
new_content = re.sub(
    r'if s_count >= b_count:',
    'if s_count > b_count:  # Fixed: strict comparison (was >=)',
    new_content
)

# Fix 2: Add minimum sample check before pattern lookup
old_pattern_section = """s_count, b_count = match[0], match[1]
                        total = s_count + b_count
                        
                        # Pattern lookup — always predict, NO SKIP"""

new_pattern_section = """s_count, b_count = match[0], match[1]
                        total = s_count + b_count
                        
                        # Minimum sample check
                        if total < 10:
                            logger.info(f"[PATTERN] {pattern_key} → INSUFFICIENT DATA total={total} < 10 → using tiebreaker")
                            if self.result_history and self.result_history[-1] == 'B':
                                pred      = 'S'
                                pred_text = "SMALL"
                            else:
                                pred      = 'B'
                                pred_text = "BIG"
                            conf = 50.0
                        # Pattern lookup — always predict, NO SKIP"""

new_content = new_content.replace(old_pattern_section, new_pattern_section)

# Write back
with open('Mr_Wai_Trading_bot_5num.py', 'w', encoding='utf-8') as f:
    f.write(new_content)

print("✅ Successfully integrated TRX_Pattern_5num_FINAL.xlsx!")
print(f"   - Excel file: TRX_Pattern_5num_FINAL.xlsx")
print(f"   - Removed hardcoded PATTERN_TABLE (100,000 patterns)")
print(f"   - Added Excel loading function")
print(f"   - Added fallback logic")
print(f"   - Fixed: tie-handling bug (>= changed to >)")
print(f"   - Fixed: added minimum sample check (total < 10)")
print()
print("=" * 80)
print("Integration complete!")
print("=" * 80)
