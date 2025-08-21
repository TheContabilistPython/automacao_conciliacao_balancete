#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analisador de estrutura HTML para relatórios fiscais
Ajuda a identificar problemas na extração de dados
"""

import os
import glob
from bs4 import BeautifulSoup

def analyze_html_structure():
    """Analisa a estrutura HTML dos relatórios fiscais"""
    
    base_dir = "C:\\relatorios_fiscal"
    
    if not os.path.exists(base_dir):
        print(f"Erro: Diretório {base_dir} não encontrado!")
        return
    
    # Busca por arquivos HTML
    html_files = glob.glob(os.path.join(base_dir, "*.htm"))
    
    if not html_files:
        print(f"Nenhum arquivo HTML encontrado em {base_dir}")
        return
    
    print(f"Encontrados {len(html_files)} arquivos HTML:")
    for i, file in enumerate(html_files):
        print(f"{i+1}. {os.path.basename(file)}")
    
    # Analisa cada arquivo
    for file in html_files:
        print(f"\n{'='*60}")
        print(f"ANALISANDO: {os.path.basename(file)}")
        print('='*60)
        
        try:
            with open(file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            soup = BeautifulSoup(content, 'html.parser')
            
            # Estatísticas básicas
            print(f"Total de células <td>: {len(soup.find_all('td'))}")
            print(f"Total de linhas <tr>: {len(soup.find_all('tr'))}")
            
            # Busca por texto relacionado ao IPI
            ipi_cells = soup.find_all('td', string=lambda text: text and 'IPI' in text)
            print(f"Células contendo 'IPI': {len(ipi_cells)}")
            
            for i, cell in enumerate(ipi_cells):
                print(f"  {i+1}: '{cell.get_text(strip=True)}'")
                
                # Mostra atributos da célula
                if cell.attrs:
                    print(f"      Atributos: {cell.attrs}")
                
                # Mostra células adjacentes na mesma linha
                row = cell.find_parent('tr')
                if row:
                    cells_in_row = row.find_all('td')
                    print(f"      Células na linha ({len(cells_in_row)} total):")
                    for j, row_cell in enumerate(cells_in_row):
                        cell_text = row_cell.get_text(strip=True)
                        if cell_text:
                            print(f"        [{j}]: '{cell_text}' {row_cell.attrs}")
            
            # Busca por valores monetários (números com vírgulas)
            monetary_cells = []
            all_cells = soup.find_all('td')
            
            for cell in all_cells:
                cell_text = cell.get_text(strip=True)
                if ',' in cell_text and cell_text.replace('.', '').replace(',', '').replace(' ', '').isdigit():
                    try:
                        # Tenta converter para número
                        value = float(cell_text.replace('.', '').replace(',', '.'))
                        if value >= 100:  # Valores significativos
                            monetary_cells.append((cell_text, value, cell.attrs))
                    except:
                        pass
            
            print(f"\nValores monetários encontrados (>= 100): {len(monetary_cells)}")
            for text, value, attrs in monetary_cells[:10]:  # Mostra até 10
                print(f"  {text} = {value:.2f} {attrs}")
                
                # Verifica se está próximo a texto relacionado ao IPI
                cell = soup.find('td', string=text)
                if cell:
                    row = cell.find_parent('tr')
                    if row:
                        row_text = row.get_text().lower()
                        if 'ipi' in row_text:
                            print(f"    ⚠️  POSSÍVEL VALOR DO IPI: {value:.2f}")
            
        except Exception as e:
            print(f"Erro ao analisar {file}: {e}")

if __name__ == "__main__":
    analyze_html_structure()
