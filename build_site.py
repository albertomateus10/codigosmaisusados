# -*- coding: utf-8 -*-
"""
Script to generate index.html and tabela_pecas_carros.html
from lista_pecas_carros.xlsx with the visual design of https://albertomateus10.github.io/temposdegarantia/
"""

import json
import zipfile
import xml.etree.ElementTree as ET

# 1. Read Excel file lista_pecas_carros.xlsx
with zipfile.ZipFile('lista_pecas_carros.xlsx') as z:
    sst = []
    if 'xl/sharedStrings.xml' in z.namelist():
        tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
            t = si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
            if t is not None:
                sst.append(t.text or '')
            else:
                text_parts = [elem.text or '' for elem in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')]
                sst.append(''.join(text_parts))

    sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    rows = []
    for row in sheet.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
        row_vals = {}
        for c in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            r = c.attrib.get('r', '')
            col = ''.join(filter(str.isalpha, r))
            t = c.attrib.get('t', '')
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else ''
            if t == 's' and val != '':
                val = sst[int(val)]
            row_vals[col] = val.strip()
        rows.append(row_vals)

items = []
car_model_counts = {}

for r in rows[1:]:
    codigo = r.get('A', '').strip()
    nome = r.get('B', '').strip()
    carros = r.get('C', '').strip()
    if codigo or nome or carros:
        car_list = [c.strip() for c in carros.splitlines() if c.strip()]
        for c in car_list:
            car_model_counts[c] = car_model_counts.get(c, 0) + 1
        items.append({
            'codigo': codigo,
            'nome': nome,
            'carros': carros,
            'carros_list': car_list,
            'num_carros': len(car_list)
        })

print(f"Loaded {len(items)} items. Found {len(car_model_counts)} distinct car models.")

sorted_cars = sorted(car_model_counts.keys())
embedded_json = json.dumps(items, ensure_ascii=False)
sorted_cars_json = json.dumps(sorted_cars, ensure_ascii=False)

html_content = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tabela de Peças por Carro - Excel Style</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: Calibri, Arial, 'Segoe UI', sans-serif;
        }}

        body {{
            background-color: #ffffff;
            color: #000000;
            padding: 15px;
        }}

        /* Top ribbon / Filter bar */
        .filter-ribbon {{
            background-color: #f3f3f3;
            border: 1px solid #d4d4d4;
            padding: 10px 15px;
            margin-bottom: 0px;
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
            font-size: 10pt;
        }}

        .filter-ribbon h2 {{
            font-size: 11pt;
            font-weight: bold;
            color: #444;
            border-right: 1px solid #ccc;
            padding-right: 15px;
            margin-right: 5px;
        }}

        .filter-group {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .filter-group label {{
            font-weight: bold;
            color: #333;
        }}

        .filter-input {{
            border: 1px solid #bbb;
            padding: 4px 8px;
            font-size: 10pt;
            outline: none;
            background-color: #ffffff;
            height: 26px;
        }}

        .filter-input:focus {{
            border-color: #ff0000;
        }}

        select.filter-input {{
            cursor: pointer;
            padding-right: 20px;
        }}

        .btn {{
            background-color: #ffffff;
            border: 1px solid #bbb;
            color: #333;
            padding: 4px 12px;
            font-size: 10pt;
            cursor: pointer;
            height: 26px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-weight: normal;
            border-radius: 2px;
            text-decoration: none;
        }}

        .btn:hover {{
            background-color: #e9e9e9;
            border-color: #999;
        }}



        /* Quick Vehicle Family Filter Ribbon */
        .quick-chips-bar {{
            background-color: #fafafa;
            border: 1px solid #d4d4d4;
            border-top: none;
            padding: 6px 15px;
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 6px;
            margin-bottom: 15px;
            font-size: 9pt;
        }}

        .quick-chips-label {{
            font-weight: bold;
            color: #555;
            margin-right: 4px;
            font-size: 9pt;
        }}

        .quick-chip {{
            padding: 2px 10px;
            border-radius: 12px;
            border: 1px solid #ccc;
            background-color: #ffffff;
            color: #333;
            cursor: pointer;
            font-size: 8.5pt;
            transition: all 0.15s ease;
            user-select: none;
        }}

        .quick-chip:hover {{
            background-color: #eaeaea;
            border-color: #999;
            transform: translateY(-1px);
        }}

        .quick-chip.active {{
            background-color: #ff0000;
            color: #ffffff;
            border-color: #cc0000;
            font-weight: bold;
        }}

        .quick-chip-count {{
            font-size: 7.5pt;
            opacity: 0.85;
            margin-left: 2px;
        }}



        /* Excel Grid Layout */
        .excel-container {{
            border: 1px solid #ccc;
            overflow-x: auto;
            background-color: #ffffff;
            position: relative;
            max-height: calc(100vh - 150px);
            overflow-y: auto;
        }}

        table {{
            width: 100%;
            min-width: 900px;
            border-collapse: collapse;
            font-size: 10pt;
            text-align: left;
            background-color: #ffffff;
        }}

        /* Header Row (Bright Red background) */
        thead {{
            position: sticky;
            top: 0;
            z-index: 10;
        }}

        th {{
            background-color: #ff0000;
            color: #ffffff;
            font-weight: bold;
            text-align: center;
            padding: 6px 12px;
            border: 1px solid #c0c0c0;
            font-size: 10pt;
            letter-spacing: 0.3px;
            white-space: nowrap;
        }}

        th.sortable {{
            cursor: pointer;
        }}

        th.sortable:hover {{
            background-color: #e00000;
        }}

        /* Rows and cells */
        td {{
            padding: 6px 12px;
            border: 1px solid #d4d4d4;
            color: #000000;
            vertical-align: middle;
        }}

        tr:hover td {{
            background-color: #f2f7fc;
        }}

        /* Left row numbers (like Excel) */
        .row-num {{
            background-color: #e4e4e4;
            color: #555555;
            text-align: center;
            font-size: 8.5pt;
            font-weight: normal;
            border: 1px solid #c0c0c0;
            width: 45px;
            min-width: 45px;
            user-select: none;
        }}

        .code-cell {{
            font-weight: bold;
            color: #000;
            font-family: Consolas, 'Segoe UI', monospace;
            font-size: 10.5pt;
            text-align: center;
        }}

        /* Car Tags / Badges */
        .car-tags-wrapper {{
            display: flex;
            flex-wrap: wrap;
            gap: 4px;
            align-items: center;
        }}

        .car-tag {{
            display: inline-block;
            background-color: #f0f4f8;
            color: #1e3a5f;
            border: 1px solid #cde0f0;
            border-radius: 3px;
            padding: 2px 7px;
            font-size: 8.5pt;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.15s ease;
            white-space: nowrap;
        }}

        .car-tag:hover {{
            background-color: #ff0000;
            color: #ffffff;
            border-color: #cc0000;
            transform: translateY(-1px);
        }}

        .car-count-badge {{
            display: inline-block;
            background-color: #ececec;
            color: #666;
            font-size: 7.5pt;
            padding: 1px 5px;
            border-radius: 8px;
            margin-left: 5px;
            font-weight: normal;
        }}


        .highlight {{
            background-color: #ffeb9c;
            color: #9c6500;
            padding: 0 2px;
            font-weight: bold;
        }}

        /* Toast message */
        .toast-container {{
            position: fixed;
            bottom: 15px;
            right: 15px;
            z-index: 1000;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}

        .toast {{
            background-color: #333;
            color: #fff;
            padding: 8px 15px;
            font-size: 9pt;
            border-radius: 2px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            opacity: 0;
            transform: translateY(20px);
            animation: slideIn 0.2s forwards, fadeOut 0.2s 2.3s forwards;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        @keyframes slideIn {{
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        @keyframes fadeOut {{
            to {{ opacity: 0; transform: translateY(-10px); }}
        }}

        .empty-state {{
            padding: 30px;
            text-align: center;
            color: #777;
            font-style: italic;
            border: 1px solid #d4d4d4;
            border-top: none;
            background-color: #ffffff;
            font-size: 10.5pt;
        }}

        @media print {{
            .filter-ribbon, .quick-chips-bar {{
                display: none !important;
            }}
            .excel-container {{
                border: none;
                max-height: none;
                overflow: visible;
            }}
            table {{
                width: 100%;
                font-size: 9pt;
            }}
            th {{
                background-color: #333 !important;
                color: #fff !important;
            }}
        }}
    </style>
</head>
<body>

    <!-- Filters Bar (Ribbon Style) -->
    <div class="filter-ribbon">
        <h2>Filtros</h2>
        <div class="filter-group">
            <label for="filterCodigo">Código:</label>
            <input type="text" id="filterCodigo" class="filter-input" placeholder="Pesquisar..." style="width: 130px;">
        </div>
        <div class="filter-group">
            <label for="filterDescricao">Descrição / Nome:</label>
            <input type="text" id="filterDescricao" class="filter-input" placeholder="Pesquisar..." style="width: 220px;">
        </div>
        <div class="filter-group">
            <label for="filterCarro">Carro:</label>
            <select id="filterCarro" class="filter-input" style="max-width: 250px;">
                <option value="">(Todos os Carros)</option>
            </select>
        </div>
        <button class="btn" id="btnClearFilters">Limpar</button>
    </div>

    <!-- Quick Vehicle Family Filter Ribbon -->
    <div class="quick-chips-bar" id="quickChipsBar">
        <span class="quick-chips-label">Atalhos por Modelo:</span>
        <button class="quick-chip active" data-family="">TODOS (101)</button>
        <button class="quick-chip" data-family="TORO">TORO (33)</button>
        <button class="quick-chip" data-family="DUCATO">DUCATO (24)</button>
        <button class="quick-chip" data-family="STRADA">STRADA (20)</button>
        <button class="quick-chip" data-family="TITANO">TITANO (20)</button>
        <button class="quick-chip" data-family="SCUDO">SCUDO (20)</button>
        <button class="quick-chip" data-family="MOBI">MOBI (16)</button>
        <button class="quick-chip" data-family="FIORINO">FIORINO (15)</button>
        <button class="quick-chip" data-family="CRONOS">CRONOS (14)</button>
        <button class="quick-chip" data-family="PULSE">PULSE (14)</button>
        <button class="quick-chip" data-family="UNO">UNO (14)</button>
        <button class="quick-chip" data-family="ARGO">ARGO (13)</button>
        <button class="quick-chip" data-family="FASTBACK">FASTBACK (9)</button>
        <button class="quick-chip" data-family="PALIO">PALIO/SIENA</button>
    </div>

    <!-- Table Sheet Container -->
    <div class="excel-container">
        <table id="itemsTable">
            <thead>
                <tr>
                    <th style="width: 45px;"></th>
                    <th class="sortable" id="thCodigo" style="width: 140px; text-align: center;">
                        CÓDIGO ↕
                    </th>
                    <th class="sortable" id="thNome" style="width: 320px; text-align: left;">
                        Descrição DE PEÇA ↕
                    </th>
                    <th style="text-align: left;">
                        CARROS EM QUE É UTILIZADA
                    </th>
                </tr>
            </thead>
            <tbody id="tableBody">
                <!-- Dynamic Excel Rows -->
            </tbody>
        </table>
    </div>

    <!-- Empty state -->
    <div id="emptyState" class="empty-state" style="display: none;">
        Nenhum registro encontrado para os filtros selecionados.
    </div>

    <!-- Toast Notification -->
    <div class="toast-container" id="toastContainer"></div>

    <script>
        // Injected data from lista_pecas_carros.xlsx
        window.ITEMS_DATA = {embedded_json};
        window.CAR_MODELS = {sorted_cars_json};

        // Application State
        const state = {{
            data: window.ITEMS_DATA,
            filteredData: [...window.ITEMS_DATA],
            filters: {{
                codigo: '',
                descricao: '',
                carro: '',
                family: ''
            }},
            sort: {{
                column: 'nome',
                direction: 'asc'
            }}
        }};

        // Elements
        const filterCodigo = document.getElementById('filterCodigo');
        const filterDescricao = document.getElementById('filterDescricao');
        const filterCarro = document.getElementById('filterCarro');
        const btnClearFilters = document.getElementById('btnClearFilters');
        const tableBody = document.getElementById('tableBody');
        const itemsTable = document.getElementById('itemsTable');
        const emptyState = document.getElementById('emptyState');
        const toastContainer = document.getElementById('toastContainer');
        const quickChipsBar = document.getElementById('quickChipsBar');
        const thCodigo = document.getElementById('thCodigo');
        const thNome = document.getElementById('thNome');

        // Normalize string (lowercase and remove diacritics)
        function normalizeStr(str) {{
            if (!str) return '';
            return str.toString()
                      .normalize('NFD')
                      .replace(/[\\u0300-\\u036f]/g, '')
                      .toLowerCase()
                      .trim();
        }}

        // Populate Car Models Dropdown
        function populateCarDropdown() {{
            window.CAR_MODELS.forEach(car => {{
                const count = state.data.filter(item => item.carros_list.includes(car)).length;
                const opt = document.createElement('option');
                opt.value = car;
                opt.textContent = car + ' (' + count + ')';
                filterCarro.appendChild(opt);
            }});
        }}

        // Toast feedback
        function showToast(message) {{
            const toast = document.createElement('div');
            toast.className = 'toast';
            toast.innerText = message;
            toastContainer.appendChild(toast);

            setTimeout(() => {{
                if (toast && toast.parentNode) {{
                    toast.parentNode.removeChild(toast);
                }}
            }}, 2500);
        }}

        // Copy text helper
        function copyText(text, msg) {{
            if (navigator.clipboard && window.isSecureContext) {{
                navigator.clipboard.writeText(text).then(() => {{
                    showToast(msg);
                }}).catch(() => {{
                    fallbackCopy(text, msg);
                }});
            }} else {{
                fallbackCopy(text, msg);
            }}
        }}

        function fallbackCopy(text, msg) {{
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.left = '-9999px';
            document.body.appendChild(ta);
            ta.focus();
            ta.select();
            try {{
                document.execCommand('copy');
                showToast(msg);
            }} catch (e) {{
                showToast('Erro ao copiar texto');
            }}
            document.body.removeChild(ta);
        }}

        // Escape HTML
        function escapeHtml(text) {{
            if (!text) return '';
            return text
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }}

        // Highlight matched text helper
        function highlightText(text, search) {{
            if (!search || !text) return escapeHtml(text);
            const normalizedText = normalizeStr(text);
            const normalizedSearch = normalizeStr(search);
            
            let result = '';
            let lastIdx = 0;
            let idx = normalizedText.indexOf(normalizedSearch);
            
            if (idx === -1) return escapeHtml(text);

            while (idx !== -1) {{
                result += escapeHtml(text.substring(lastIdx, idx)) + 
                          '<mark class="highlight">' + escapeHtml(text.substring(idx, idx + search.length)) + '</mark>';
                lastIdx = idx + search.length;
                idx = normalizedText.indexOf(normalizedSearch, lastIdx);
            }}
            result += escapeHtml(text.substring(lastIdx));
            return result;
        }}

        // Apply filters
        function applyFilters() {{
            const filterCodVal = normalizeStr(state.filters.codigo);
            const filterDescVal = normalizeStr(state.filters.descricao);
            const filterCarroVal = state.filters.carro;
            const filterFamVal = normalizeStr(state.filters.family);

            state.filteredData = state.data.filter(item => {{
                const matchCode = !filterCodVal || normalizeStr(item.codigo).includes(filterCodVal);
                const matchDesc = !filterDescVal || normalizeStr(item.nome).includes(filterDescVal);
                const matchCarro = !filterCarroVal || item.carros_list.includes(filterCarroVal);
                const matchFam = !filterFamVal || item.carros_list.some(c => normalizeStr(c).includes(filterFamVal));

                return matchCode && matchDesc && matchCarro && matchFam;
            }});

            // Sort
            sortData();

            renderTable();
        }}

        // Sort data
        function sortData() {{
            const col = state.sort.column;
            const dir = state.sort.direction === 'asc' ? 1 : -1;

            state.filteredData.sort((a, b) => {{
                let valA = a[col] || '';
                let valB = b[col] || '';
                return valA.localeCompare(valB, 'pt-BR', {{ sensitivity: 'base' }}) * dir;
            }});
        }}

        // Render table
        function renderTable() {{
            const totalItems = state.filteredData.length;

            if (totalItems === 0) {{
                tableBody.innerHTML = '';
                itemsTable.style.display = 'none';
                emptyState.style.display = 'block';
                return;
            }}

            emptyState.style.display = 'none';
            itemsTable.style.display = 'table';

            let html = '';
            state.filteredData.forEach((item, idx) => {{
                const rowNum = idx + 2; // Excel index starting at 2
                const escCode = escapeHtml(item.codigo);
                const escDesc = escapeHtml(item.nome);

                const highlightedCode = highlightText(item.codigo, state.filters.codigo);
                const highlightedDesc = highlightText(item.nome, state.filters.descricao);

                // Build interactive car tags
                let carTagsHtml = '';
                item.carros_list.forEach(car => {{
                    const isMatchedCar = (state.filters.carro === car) || 
                                         (state.filters.family && normalizeStr(car).includes(normalizeStr(state.filters.family)));
                    const tagStyle = isMatchedCar ? 'background-color: #ffeb9c; border-color: #e0b84c; color: #7a4f00; font-weight: bold;' : '';
                    carTagsHtml += `<span class="car-tag" style="${{tagStyle}}" onclick="filterByCarTag(event, '${{escapeHtml(car)}}')">${{escapeHtml(car)}}</span>`;
                }});

                html += `
                <tr>
                    <td class="row-num">${{rowNum}}</td>
                    <td class="code-cell">${{highlightedCode}}</td>
                    <td>
                        <strong>${{highlightedDesc}}</strong>
                        <span class="car-count-badge">${{item.num_carros}} ${{item.num_carros === 1 ? 'carro' : 'carros'}}</span>
                    </td>
                    <td>
                        <div class="car-tags-wrapper">
                            ${{carTagsHtml}}
                        </div>
                    </td>
                </tr>
                `;
            }});

            tableBody.innerHTML = html;
        }}

        // Copy actions
        window.handleCopyCode = function(event, code) {{
            if (event) event.stopPropagation();
            copyText(code, "Código copiado: " + code);
        }};

        window.filterByCarTag = function(event, carName) {{
            if (event) event.stopPropagation();
            filterCarro.value = carName;
            state.filters.carro = carName;

            document.querySelectorAll('.quick-chip').forEach(c => c.classList.remove('active'));
            state.filters.family = '';

            applyFilters();
            showToast("Filtrando por: " + carName);
        }};

        // Listeners for filters
        filterCodigo.addEventListener('input', (e) => {{
            state.filters.codigo = e.target.value;
            applyFilters();
        }});

        filterDescricao.addEventListener('input', (e) => {{
            state.filters.descricao = e.target.value;
            applyFilters();
        }});

        filterCarro.addEventListener('change', (e) => {{
            state.filters.carro = e.target.value;
            document.querySelectorAll('.quick-chip').forEach(c => c.classList.remove('active'));
            state.filters.family = '';
            applyFilters();
        }});

        // Quick family chips
        quickChipsBar.addEventListener('click', (e) => {{
            const chip = e.target.closest('.quick-chip');
            if (!chip) return;

            document.querySelectorAll('.quick-chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');

            const fam = chip.getAttribute('data-family');
            state.filters.family = fam;

            filterCarro.value = '';
            state.filters.carro = '';

            applyFilters();
            if (fam) {{
                showToast("Filtrado pela família: " + fam);
            }} else {{
                showToast("Mostrando todos os carros");
            }}
        }});

        // Clear button
        btnClearFilters.addEventListener('click', () => {{
            filterCodigo.value = '';
            filterDescricao.value = '';
            filterCarro.value = '';

            state.filters.codigo = '';
            state.filters.descricao = '';
            state.filters.carro = '';
            state.filters.family = '';

            document.querySelectorAll('.quick-chip').forEach(c => c.classList.remove('active'));
            const defaultChip = document.querySelector('.quick-chip[data-family=""]');
            if (defaultChip) defaultChip.classList.add('active');

            applyFilters();
            showToast('Filtros redefinidos');
        }});

        // Esc shortcut to clear filters
        document.addEventListener('keydown', (e) => {{
            if (e.key === 'Escape') {{
                btnClearFilters.click();
            }}
        }});

        // Sort headers
        thCodigo.addEventListener('click', () => {{
            if (state.sort.column === 'codigo') {{
                state.sort.direction = state.sort.direction === 'asc' ? 'desc' : 'asc';
            }} else {{
                state.sort.column = 'codigo';
                state.sort.direction = 'asc';
            }}
            thCodigo.innerText = 'CÓDIGO ' + (state.sort.direction === 'asc' ? '▲' : '▼');
            thNome.innerText = 'Descrição DE PEÇA ↕';
            applyFilters();
        }});

        thNome.addEventListener('click', () => {{
            if (state.sort.column === 'nome') {{
                state.sort.direction = state.sort.direction === 'asc' ? 'desc' : 'asc';
            }} else {{
                state.sort.column = 'nome';
                state.sort.direction = 'asc';
            }}
            thNome.innerText = 'Descrição DE PEÇA ' + (state.sort.direction === 'asc' ? '▲' : '▼');
            thCodigo.innerText = 'CÓDIGO ↕';
            applyFilters();
        }});



        // Initialize Data
        populateCarDropdown();
        applyFilters();
    </script>
</body>
</html>
'''

# Write to index.html and tabela_pecas_carros.html
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

with open('tabela_pecas_carros.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("Generated index.html and tabela_pecas_carros.html successfully.")
