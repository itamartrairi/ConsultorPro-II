import re

with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

print(f"Total lines in App.tsx: {len(lines)}")

# Let's find component line ranges
def find_range(start_pat, end_pat):
    start_idx = None
    end_idx = None
    for i, l in enumerate(lines):
        if start_idx is None and re.search(start_pat, l):
            start_idx = i
        elif start_idx is not None and end_pat and re.search(end_pat, l):
            end_idx = i
            break
    return start_idx, end_idx

ranges = {
    'DadosConsultoriaView': find_range(r'const DadosConsultoriaView\s*=', r'const CronogramaView\s*='),
    'RelatorioView': find_range(r'const RelatorioView\s*=', r'const GestaoPlanoView\s*='),
    'GestaoPlanoView': find_range(r'const GestaoPlanoView\s*=', r'const LandingPage\s*='),
    'LandingPage': find_range(r'const LandingPage\s*=', r'const CheckoutPage\s*='),
    'CheckoutPage': find_range(r'const CheckoutPage\s*=', r'const LogoSelector\s*='),
    'SettingsView': find_range(r'const LogoSelector\s*=', r'const LicenseManagementView\s*='),
    'LicenseManagementView': find_range(r'const LicenseManagementView\s*=', r'const LicensingDataView\s*='),
    'LicensingDataView': find_range(r'const LicensingDataView\s*=', r'const HomeView\s*='),
    'HomeView': find_range(r'const HomeView\s*=', r'const HARDCODED_GEMINI_API_KEY\s*=')
}

for name, (s, e) in ranges.items():
    print(f"{name}: lines {s+1 if s is not None else 'None'} to {e+1 if e is not None else 'None'} (length: {(e - s) if s and e else 0})")
