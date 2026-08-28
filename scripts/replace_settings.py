with open('src/App.tsx', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# lines[8389:9259] is LogoSelector and SettingsView
del lines[8389:9259]

import_stmt = "import SettingsView from './components/SettingsView';\n"
for i, l in enumerate(lines):
    if "import { ConnectionStatus }" in l:
        lines.insert(i + 1, import_stmt)
        break

with open('src/App.tsx', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("SettingsView replaced!")
