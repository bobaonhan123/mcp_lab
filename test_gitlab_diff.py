"""Test GitLab-style diff in combined Excel."""
from src.helpers.code_ast import capture_function
from src.helpers.search import search_in_folder
from src.helpers.combined_excel import write_combined_excel

# 1. Code blocks
code_blocks = [capture_function('src/helpers/code.py', 'detect_language')]
print(f'Captured {len(code_blocks)} code blocks')

# 2. Diff - old vs new (realistic example)
old_code = '''def calculate_total(items):
    total = 0
    for item in items:
        total += item.price
    return total

def format_currency(amount):
    return f"${amount:.2f}"'''

new_code = '''def calculate_total(items, tax_rate=0.1):
    subtotal = 0
    for item in items:
        subtotal += item.price * item.quantity
    tax = subtotal * tax_rate
    return subtotal + tax

def format_currency(amount, symbol="$"):
    return f"{symbol}{amount:,.2f}"

def apply_discount(total, discount_percent):
    return total * (1 - discount_percent / 100)'''

# 3. Search
search = search_in_folder('src/helpers', 'Monokai', '*.py')
print(f'Found {search.total_matches} matches')

# 4. PlantUML
puml = '''@startuml
skinparam backgroundColor #272822
participant Client
participant Server
participant Database

Client -> Server: Request
Server -> Database: Query
Database --> Server: Results
Server --> Client: Response
@enduml'''

# Write combined Excel
path = write_combined_excel(
    output_path='output/gitlab_diff_test.xlsx',
    code_blocks=code_blocks,
    diff_old=old_code,
    diff_new=new_code,
    search_summary=search,
    puml_code=puml
)
print(f'✅ Created: {path}')
print('📊 Sheets: Code, Diff (GitLab-style), Search, Diagram')
