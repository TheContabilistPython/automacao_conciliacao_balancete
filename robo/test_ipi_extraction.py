#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de teste para validação da extração do IPI
Testa diferentes variações de nome de arquivo e estruturas HTML
"""

import os
from bs4 import BeautifulSoup

def test_ipi_extraction():
    """Testa a extração do valor do IPI com diferentes abordagens"""
    
    # Diretório base
    base_dir = "C:\\relatorios_fiscal"
    
    # Variações possíveis do nome do arquivo
    filename_patterns = [
        "Empresa 8801 - DOMAZZI S.A. - Relatório apuração IPI -  Sequência 22 - Ordem 2 01-05-2025 31-05-2025.htm",
        "Empresa 8801 - DOMAZZI S.A. - Relatório apuração IPI - Sequência 22 - Ordem 2 01-05-2025 31-05-2025.htm",
        "Empresa 8801 - DOMAZZI S.A. - Relatório apuração IPI -   Sequência 22 - Ordem 2 01-05-2025 31-05-2025.htm"
    ]
    
    for pattern in filename_patterns:
        filepath = os.path.join(base_dir, pattern)
        print(f"\nTestando arquivo: {pattern}")
        
        if os.path.exists(filepath):
            print("✓ Arquivo encontrado!")
            try:
                with open(filepath, 'r', encoding='utf-8') as file:
                    content = file.read()
                
                # Verifica conteúdo básico
                if 'IPI' in content:
                    print("✓ Contém texto 'IPI'")
                else:
                    print("✗ NÃO contém texto 'IPI'")
                
                if 'recolher' in content:
                    print("✓ Contém texto 'recolher'")
                else:
                    print("✗ NÃO contém texto 'recolher'")
                
                # Testa extração
                soup = BeautifulSoup(content, 'html.parser')
                extracted_value = extract_ipi_value(soup)
                
                if extracted_value is not None:
                    print(f"✓ Valor extraído: {extracted_value:.2f}".replace('.', ','))
                else:
                    print("✗ Valor não extraído")
                
            except Exception as e:
                print(f"✗ Erro ao processar arquivo: {e}")
        else:
            print("✗ Arquivo não encontrado")

def extract_ipi_value(soup):
    """Extrai o valor do IPI usando múltiplas abordagens"""
    
    # Abordagem 1: Busca pela estrutura específica
    text_cell = soup.find('td', {'colspan': '7', 'class': 's6'}, 
                         string=lambda text: text and "Valor do saldo devedor do IPI a recolher" in text)
    
    if text_cell:
        print("  Método 1: Encontrou célula de texto específica")
        row = text_cell.find_parent('tr')
        cells = row.find_all('td')
        
        for i, cell in enumerate(cells):
            cell_text = cell.get_text(strip=True)
            print(f"    Célula {i}: '{cell_text}'")
            
            if ',' in cell_text or ('.' in cell_text and len(cell_text.replace('.', '').replace(',', '')) > 2):
                try:
                    test_value = float(cell_text.replace('.', '').replace(',', '.'))
                    if test_value >= 100:
                        print(f"  ✓ Valor encontrado pelo método 1: {test_value}")
                        return test_value
                except ValueError:
                    continue
    
    # Abordagem 2: Busca mais ampla
    print("  Método 2: Busca ampla por 'IPI a recolher'")
    ipi_cells = soup.find_all('td', string=lambda text: text and "IPI a recolher" in text)
    
    for cell in ipi_cells:
        print(f"    Encontrou célula: '{cell.get_text(strip=True)}'")
        row = cell.find_parent('tr')
        cells = row.find_all('td')
        
        for value_cell in cells:
            value_text = value_cell.get_text(strip=True)
            if ',' in value_text or ('.' in value_text and len(value_text.replace('.', '').replace(',', '')) > 2):
                try:
                    test_value = float(value_text.replace('.', '').replace(',', '.'))
                    if test_value >= 100:
                        print(f"  ✓ Valor encontrado pelo método 2: {test_value}")
                        return test_value
                except ValueError:
                    continue
    
    # Abordagem 3: Busca por padrões monetários
    print("  Método 3: Busca por valores monetários próximos a 'IPI'")
    all_cells = soup.find_all('td')
    monetary_values = []
    
    for cell in all_cells:
        cell_text = cell.get_text(strip=True)
        if ',' in cell_text and '.' in cell_text:
            try:
                test_value = float(cell_text.replace('.', '').replace(',', '.'))
                if test_value >= 100:
                    row = cell.find_parent('tr')
                    row_text = row.get_text().lower()
                    if 'ipi' in row_text and 'recolher' in row_text:
                        print(f"    Valor candidato: {test_value} na linha: {row_text[:100]}...")
                        monetary_values.append(test_value)
            except ValueError:
                continue
    
    if monetary_values:
        max_value = max(monetary_values)
        print(f"  ✓ Valor encontrado pelo método 3: {max_value}")
        return max_value
    
    print("  ✗ Nenhum valor encontrado")
    return None

if __name__ == "__main__":
    test_ipi_extraction()
