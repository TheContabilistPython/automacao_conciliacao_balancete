#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Teste específico para validar a extração do valor do IPI a recolher
do arquivo HTML real fornecido.
"""

from bs4 import BeautifulSoup
import os

def testar_extracao_ipi():
    # Caminho do arquivo HTML real
    caminho_html = r"c:\relatorios_fiscal\Empresa 8801 - DOMAZZI S.A. - Relatório apuração IPI -  Sequência 22 - Ordem 2 01-05-2025 31-05-2025.htm"
    
    print("=== TESTE DE EXTRAÇÃO DO IPI ===")
    print(f"Arquivo: {os.path.basename(caminho_html)}")
    
    if not os.path.exists(caminho_html):
        print(f"ERRO: Arquivo não encontrado: {caminho_html}")
        return
    
    try:
        with open(caminho_html, 'r', encoding='utf-8') as file:
            content = file.read()
        
        soup = BeautifulSoup(content, 'html.parser')
        
        # Busca pela estrutura específica
        text_cell = soup.find('td', {'colspan': '7', 'class': 's6'}, 
                             string=lambda text: text and "Valor do saldo devedor do IPI a recolher" in text)
        
        if text_cell:
            print("✓ Texto 'Valor do saldo devedor do IPI a recolher' encontrado")
            
            # Busca a linha que contém o texto
            row = text_cell.find_parent('tr')
            print(f"✓ Linha encontrada com {len(row.find_all('td'))} células")
            
            # Debug: mostra todas as células da linha
            cells = row.find_all('td')
            print("\nCélulas da linha:")
            for i, cell in enumerate(cells):
                colspan = cell.get('colspan', '1')
                class_attr = cell.get('class', [])
                text = cell.get_text(strip=True)
                print(f"  {i+1}. colspan={colspan}, class={class_attr}, texto='{text}'")
            
            # Busca todas as células com colspan="2" e class="s5" nesta linha
            value_cells = row.find_all('td', {'colspan': '2', 'class': 's5'})
            print(f"\n✓ Encontradas {len(value_cells)} células com colspan='2' e class='s5'")
            
            for i, cell in enumerate(value_cells):
                text = cell.get_text(strip=True)
                print(f"  Célula {i+1}: '{text}'")
            
            if value_cells:
                # Pega a última célula
                value_cell = value_cells[-1]
                ipi_value = value_cell.get_text(strip=True)
                print(f"\n✓ Última célula selecionada: '{ipi_value}'")
                
                # Tenta converter
                try:
                    test_value = float(ipi_value.replace('.', '').replace(',', '.'))
                    if test_value < 100:
                        print(f"⚠ Valor {ipi_value} parece ser código do campo (< 100)")
                    else:
                        print(f"✅ SUCESSO: IPI a recolher = R$ {test_value:,.2f}")
                except ValueError as e:
                    print(f"❌ ERRO na conversão: {e}")
            else:
                print("❌ Nenhuma célula de valor encontrada")
        else:
            print("❌ Texto 'Valor do saldo devedor do IPI a recolher' NÃO encontrado")
            
            # Debug: mostra todas as células com class s6
            s6_cells = soup.find_all('td', {'class': 's6'})
            print(f"\nDebug: Encontradas {len(s6_cells)} células com class='s6':")
            for cell in s6_cells[:10]:  # Mostra apenas as primeiras 10
                text = cell.get_text(strip=True)
                if 'IPI' in text:
                    print(f"  - '{text}'")
                    
    except Exception as e:
        print(f"❌ ERRO: {e}")

if __name__ == "__main__":
    testar_extracao_ipi()
