# worker_unimed.py — versão com debug melhorado
# Saída: resultado_unimed.xlsx com colunas: cpf | plano | categoria | timestamp

import os
import re
import time
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

# ===================== CONFIG =====================
URL = "https://unimed.coop.br/site/guia-medico#/"
HEADLESS = False           # False = mostra o navegador
TIMEOUT = 30000            # 30s
ESPERA_APOS_DIGITAR = 2.5  # espera após digitar CPF (aumentado)
PAUSA_ENTRE_CPFS = 1.0     # pausa entre CPFs
DEBUG = False              # Modo debug (mude para True se precisar investigar)

# Aceita automaticamente 'input_clientes.xlsx' OU 'input_clientes.xlsx.xlsx'
CANDIDATOS_INPUT = ["input_clientes.xlsx", "input_clientes.xlsx.xlsx"]
OUTPUT_ARQ = "resultado_unimed.xlsx"
# ==================================================

def descobrir_input() -> str:
    for nome in CANDIDATOS_INPUT:
        if os.path.exists(nome):
            return nome
    raise FileNotFoundError(
        "Não encontrei 'input_clientes.xlsx' (nem 'input_clientes.xlsx.xlsx') na pasta do projeto."
    )

def somente_digitos(s: str) -> str:
    return re.sub(r"\D", "", (s or ""))

def carregar_cpfs(caminho_excel: str) -> list[str]:
    # CRÍTICO: dtype=str para preservar zeros à esquerda
    df = pd.read_excel(caminho_excel, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    if "cpf" not in df.columns:
        raise ValueError("A planilha precisa ter uma coluna chamada 'CPF'.")
    
    cpfs = []
    for v in df["cpf"].tolist():
        if pd.isna(v) or v == '':
            continue
        # Remove tudo que não é dígito
        dig = somente_digitos(str(v))
        if not dig:
            continue
        # Garante exatamente 11 dígitos
        if len(dig) < 11:
            dig = dig.zfill(11)  # Completa com zeros à esquerda
        elif len(dig) > 11:
            dig = dig[-11:]  # Pega últimos 11
        cpfs.append(dig)
    return cpfs

# --------- Validação de CPF ----------
def validar_cpf(c: str) -> bool:
    c = somente_digitos(c)
    if len(c) != 11 or c == c[0] * 11:
        return False

    def calc_dv(cpf_parcial: str, pesos: list[int]) -> int:
        soma = sum(int(dig) * peso for dig, peso in zip(cpf_parcial, pesos))
        resto = soma % 11
        return 0 if resto < 2 else 11 - resto

    dv1 = calc_dv(c[:9], list(range(10, 1, -1)))
    dv2 = calc_dv(c[:9] + str(dv1), list(range(11, 1, -1)))
    return c[-2:] == f"{dv1}{dv2}"

def clicar_busca_detalhada(page):
    if DEBUG:
        print("   [DEBUG] Tentando clicar em 'Busca detalhada'...")
    candidatos = [
        page.get_by_role("tab", name="Busca detalhada", exact=False),
        page.locator("text=Busca detalhada").first,
        page.locator("button:has-text('Busca detalhada')").first,
        page.locator("[role='tab']:has-text('Busca detalhada')").first,
    ]
    for i, c in enumerate(candidatos):
        try:
            c.wait_for(state="visible", timeout=3000)
            c.click()
            page.wait_for_timeout(500)
            if DEBUG:
                print(f"   [DEBUG] Clicou com sucesso (método {i+1})")
            return True
        except Exception as e:
            if DEBUG:
                print(f"   [DEBUG] Método {i+1} falhou: {e}")
            continue
    return False

def abrir_ver_mais_filtros(page):
    if DEBUG:
        print("   [DEBUG] Tentando abrir 'Ver mais filtros'...")
    try:
        btn = page.locator("text=Ver mais filtros").first
        if btn.is_visible(timeout=2000):
            btn.click()
            page.wait_for_timeout(500)
            if DEBUG:
                print("   [DEBUG] 'Ver mais filtros' aberto")
            return True
    except Exception as e:
        if DEBUG:
            print(f"   [DEBUG] 'Ver mais filtros' não encontrado: {e}")
    return False

def localizar_campo_cpf(page):
    if DEBUG:
        print("   [DEBUG] Localizando campo CPF...")
    candidatos = [
        ("placeholder 000.000.000-00", page.get_by_placeholder("000.000.000-00", exact=False)),
        ("placeholder CPF", page.get_by_placeholder("CPF", exact=False)),
        ("input com placeholder 000", page.locator("input[placeholder*='000.000.000']").first),
        ("input com placeholder CPF", page.locator("input[placeholder*='CPF']").first),
        ("input type text", page.locator("input[type='text']").first),
    ]
    for nome, c in candidatos:
        try:
            c.wait_for(state="visible", timeout=3000)
            if DEBUG:
                print(f"   [DEBUG] Campo encontrado: {nome}")
            return c
        except Exception as e:
            if DEBUG:
                print(f"   [DEBUG] {nome} não encontrado: {e}")
            continue
    raise RuntimeError("Campo CPF não encontrado.")

def limpar_campo_cpf(page, campo):
    """Limpa o campo CPF usando múltiplos métodos"""
    if DEBUG:
        print("   [DEBUG] Limpando campo CPF...")
    
    # Método 1: Botão "Limpar dados"
    try:
        btn = page.locator("text=Limpar dados").first
        if btn.is_visible(timeout=1000):
            btn.click()
            page.wait_for_timeout(300)
            if DEBUG:
                print("   [DEBUG] Limpou via botão 'Limpar dados'")
            return True
    except Exception:
        pass
    
    # Método 2: Selecionar tudo e deletar
    try:
        campo.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.wait_for_timeout(200)
        if DEBUG:
            print("   [DEBUG] Limpou via Ctrl+A + Backspace")
        return True
    except Exception:
        pass
    
    # Método 3: Clear
    try:
        campo.clear()
        page.wait_for_timeout(200)
        if DEBUG:
            print("   [DEBUG] Limpou via clear()")
        return True
    except Exception:
        pass
    
    return False

def preencher_cpf_com_multiplas_estrategias(page, campo, cpf_mask):
    """Tenta preencher o CPF usando diferentes métodos"""
    if DEBUG:
        print(f"   [DEBUG] Preenchendo CPF: {cpf_mask}")
    
    # Método 1: Fill simples
    try:
        campo.fill(cpf_mask)
        page.wait_for_timeout(300)
        valor = campo.input_value()
        if DEBUG:
            print(f"   [DEBUG] Método fill() - Valor no campo: '{valor}'")
        if somente_digitos(valor) == somente_digitos(cpf_mask):
            campo.press("Tab")
            return True
    except Exception as e:
        if DEBUG:
            print(f"   [DEBUG] Método fill() falhou: {e}")
    
    # Método 2: Type com delay
    try:
        campo.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        campo.type(cpf_mask, delay=100)
        page.wait_for_timeout(300)
        valor = campo.input_value()
        if DEBUG:
            print(f"   [DEBUG] Método type() - Valor no campo: '{valor}'")
        if somente_digitos(valor) == somente_digitos(cpf_mask):
            campo.press("Tab")
            return True
    except Exception as e:
        if DEBUG:
            print(f"   [DEBUG] Método type() falhou: {e}")
    
    # Método 3: Press sequencial
    try:
        campo.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        for char in cpf_mask:
            page.keyboard.press(char)
            page.wait_for_timeout(50)
        page.wait_for_timeout(300)
        valor = campo.input_value()
        if DEBUG:
            print(f"   [DEBUG] Método press() - Valor no campo: '{valor}'")
        if somente_digitos(valor) == somente_digitos(cpf_mask):
            campo.press("Tab")
            return True
    except Exception as e:
        if DEBUG:
            print(f"   [DEBUG] Método press() falhou: {e}")
    
    # Método 4: JavaScript direto
    try:
        campo.evaluate(f"el => {{ el.value = '{cpf_mask}'; el.dispatchEvent(new Event('input', {{ bubbles: true }})); el.dispatchEvent(new Event('change', {{ bubbles: true }})); }}")
        page.wait_for_timeout(300)
        valor = campo.input_value()
        if DEBUG:
            print(f"   [DEBUG] Método JavaScript - Valor no campo: '{valor}'")
        campo.press("Tab")
        return True
    except Exception as e:
        if DEBUG:
            print(f"   [DEBUG] Método JavaScript falhou: {e}")
    
    return False

def verificar_resultado(page) -> tuple[str, str]:
    """Verifica se há resultado ou mensagem de erro"""
    if DEBUG:
        print("   [DEBUG] Verificando resultado...")
    
    page.wait_for_timeout(2000)  # Aguardar processamento (aumentado)
    
    # 1. Verificar mensagens de erro PRIMEIRO
    mensagens_erro = [
        "não foi possível localizar",
        "não encontrado",
        "dados não encontrados",
        "CPF não encontrado",
        "nenhum resultado",
        "não localizado",
    ]
    
    for msg in mensagens_erro:
        try:
            if page.locator(f"text=/{msg}/i").first.is_visible(timeout=500):
                if DEBUG:
                    print(f"   [DEBUG] Mensagem de erro detectada: {msg}")
                return ("NÃO ENCONTRADO", "")
        except Exception:
            pass
    
    # 2. Buscar informações de plano - MÚLTIPLAS ESTRATÉGIAS
    plano = ""
    categoria = ""
    
    # Estratégia A: Texto abaixo do campo CPF que contém UNIMED
    try:
        xpath = "//input[contains(@placeholder,'000.000.000')]/following::*[contains(text(),'UNIMED')]"
        elemento = page.locator(f"xpath={xpath}").first
        if elemento.is_visible(timeout=1000):
            plano = elemento.inner_text().strip()
            if DEBUG:
                print(f"   [DEBUG] Plano encontrado (xpath): {plano}")
    except Exception:
        pass
    
    # Estratégia B: Qualquer elemento visível com UNIMED
    if not plano:
        try:
            elementos = page.locator("text=/UNIMED/i").all()
            for el in elementos[:5]:  # Limitar a 5 primeiros
                try:
                    if el.is_visible():
                        texto = el.inner_text().strip()
                        if texto and len(texto) > 5 and "selecione" not in texto.lower():
                            plano = texto
                            if DEBUG:
                                print(f"   [DEBUG] Plano encontrado (text): {plano}")
                            break
                except Exception:
                    pass
        except Exception:
            pass
    
    # Estratégia C: Campo/label "Plano" ou "Categoria"
    try:
        for label_text in ["Plano", "Categoria", "Produto"]:
            try:
                label = page.get_by_label(label_text, exact=False)
                if label.is_visible(timeout=1000):
                    texto = label.inner_text().strip()
                    if texto and not re.search(r"selecione|escolha", texto, re.I):
                        if not categoria:
                            categoria = texto
                            if DEBUG:
                                print(f"   [DEBUG] Categoria encontrada ({label_text}): {categoria}")
                        break
            except Exception:
                pass
    except Exception:
        pass
    
    # Estratégia D: Buscar em divs/spans próximos ao campo CPF
    if not plano and not categoria:
        try:
            # Pegar todos os textos visíveis após o campo CPF
            elementos = page.locator("xpath=//input[contains(@placeholder,'000.000.000')]/following::div | //input[contains(@placeholder,'000.000.000')]/following::span").all()
            for el in elementos[:10]:
                try:
                    if el.is_visible():
                        texto = el.inner_text().strip()
                        if texto and len(texto) > 10 and ("UNIMED" in texto.upper() or "PLANO" in texto.upper()):
                            if not plano:
                                plano = texto
                                if DEBUG:
                                    print(f"   [DEBUG] Plano encontrado (div/span): {plano}")
                            break
                except Exception:
                    pass
        except Exception:
            pass
    
    # 3. Se não encontrou nada, tentar capturar qualquer texto relevante
    if not plano and not categoria:
        try:
            # Pegar screenshot da área de resultado para debug
            if DEBUG:
                timestamp = int(time.time())
                page.screenshot(path=f"debug_resultado_{timestamp}.png")
                print(f"   [DEBUG] Screenshot salvo: debug_resultado_{timestamp}.png")
                
                # Tentar capturar HTML da página para análise
                try:
                    html = page.content()
                    with open(f"debug_html_{timestamp}.html", "w", encoding="utf-8") as f:
                        f.write(html)
                    print(f"   [DEBUG] HTML salvo: debug_html_{timestamp}.html")
                except Exception:
                    pass
        except Exception:
            pass
    
    # 4. Verificar se há algum resultado positivo (mesmo que não tenha conseguido extrair)
    # Se não tem mensagem de erro E há mudanças na página, pode ter resultado
    if not plano and not categoria:
        try:
            # Verificar se apareceu algum card, modal ou seção de resultado
            indicadores_resultado = [
                page.locator("[class*='resultado']").first,
                page.locator("[class*='card']").first,
                page.locator("[class*='info']").first,
            ]
            for ind in indicadores_resultado:
                try:
                    if ind.is_visible(timeout=500):
                        if DEBUG:
                            print("   [DEBUG] Indicador de resultado detectado, mas dados não extraídos")
                        return ("DADOS ENCONTRADOS (não extraídos)", "Verificar manualmente")
                except Exception:
                    pass
        except Exception:
            pass
        
        return ("NÃO ENCONTRADO", "")
    
    return (plano or "N/A", categoria or "N/A")

def consultar_um_cpf(page, cpf: str) -> dict:
    """Consulta um CPF e retorna plano/categoria"""
    # Ir para busca detalhada
    clicar_busca_detalhada(page)
    abrir_ver_mais_filtros(page)
    
    # Localizar campo
    campo = localizar_campo_cpf(page)
    
    # Limpar campo primeiro
    limpar_campo_cpf(page, campo)
    
    # Formatar CPF com máscara
    if len(cpf) != 11:
        raise ValueError(f"CPF deve ter 11 dígitos, recebeu {len(cpf)}")
    cpf_mask = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}"
    
    # Preencher CPF
    sucesso = preencher_cpf_com_multiplas_estrategias(page, campo, cpf_mask)
    if not sucesso and DEBUG:
        print("   [DEBUG] ⚠️ ATENÇÃO: Nenhum método de preenchimento foi bem-sucedido!")
    
    # Aguardar processamento
    page.wait_for_timeout(int(ESPERA_APOS_DIGITAR * 1000))
    
    # Verificar resultado
    plano, categoria = verificar_resultado(page)
    
    return {"plano": plano, "categoria": categoria}

# ---------------------- main -----------------------------

def main():
    INPUT_ARQ = descobrir_input()
    cpfs = carregar_cpfs(INPUT_ARQ)
    linhas = []
    
    print(f"\n{'='*60}")
    print(f"CONSULTA UNIMED - {len(cpfs)} CPFs")
    print(f"Modo DEBUG: {'ATIVO ⚠️' if DEBUG else 'DESATIVADO'}")
    print(f"{'='*60}\n")
    
    # Estatísticas
    stats = {
        "encontrados": 0,
        "nao_encontrados": 0,
        "erros": 0,
        "cpfs_invalidos": 0
    }

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=HEADLESS)
        ctx = browser.new_context(viewport={"width": 1280, "height": 720})
        page = ctx.new_page()
        
        print("🌐 Abrindo site Unimed...")
        page.goto(URL, timeout=TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)  # Aguardar carregamento inicial
        print("✓ Site carregado\n")

        for i, cpf in enumerate(cpfs, start=1):
            # Garante exatamente 11 dígitos
            cpf = somente_digitos(cpf)
            if len(cpf) > 11:
                cpf = cpf[-11:]  # Pega últimos 11
            elif len(cpf) < 11:
                cpf = cpf.zfill(11)  # Completa com zeros à esquerda
            
            cpf_formatado = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}"
            print(f"[{i}/{len(cpfs)}] CPF: {cpf_formatado}", end="")

            # Validar CPF
            cpf_valido = validar_cpf(cpf) if len(cpf) == 11 else False
            if not cpf_valido:
                print(" ⚠️ CPF inválido", end="")
                stats["cpfs_invalidos"] += 1

            # Consultar
            try:
                r = consultar_um_cpf(page, cpf)
                
                # Classificar resultado
                if "NÃO ENCONTRADO" in r['plano']:
                    print(" → ❌ Não encontrado")
                    stats["nao_encontrados"] += 1
                elif "ERRO" in r['plano']:
                    print(f" → ⚠️ Erro: {r['categoria'][:50]}")
                    stats["erros"] += 1
                else:
                    print(f" → ✅ ENCONTRADO!")
                    print(f"    📋 Plano: {r['plano']}")
                    if r['categoria'] and r['categoria'] != 'N/A':
                        print(f"    🏷️  Categoria: {r['categoria']}")
                    stats["encontrados"] += 1
                    
            except Exception as e:
                print(f" → ⚠️ Erro: {str(e)[:50]}")
                r = {"plano": "ERRO", "categoria": str(e)[:100]}
                stats["erros"] += 1
            
            # Salvar resultado
            r["cpf"] = cpf
            r["timestamp"] = datetime.now().isoformat(timespec="seconds")
            linhas.append(r)
            
            # Pausa entre consultas
            time.sleep(PAUSA_ENTRE_CPFS)
            
            # A cada 50 CPFs, salvar backup e mostrar estatísticas
            if i % 50 == 0:
                df_temp = pd.DataFrame(linhas, columns=["cpf", "plano", "categoria", "timestamp"])
                df_temp.to_excel(f"backup_{OUTPUT_ARQ}", index=False)
                print(f"\n{'─'*60}")
                print(f"💾 Backup salvo | Progresso: {i}/{len(cpfs)} ({i*100//len(cpfs)}%)")
                print(f"✅ Encontrados: {stats['encontrados']} | ❌ Não encontrados: {stats['nao_encontrados']} | ⚠️ Erros: {stats['erros']}")
                print(f"{'─'*60}\n")

        ctx.close()
        browser.close()

    # Salvar resultados finais
    df = pd.DataFrame(linhas, columns=["cpf", "plano", "categoria", "timestamp"])
    df.to_excel(OUTPUT_ARQ, index=False)
    
    # Relatório final
    print(f"\n{'='*60}")
    print(f"✅ PROCESSAMENTO CONCLUÍDO!")
    print(f"{'='*60}")
    print(f"📊 ESTATÍSTICAS FINAIS:")
    print(f"   Total processado: {len(cpfs)} CPFs")
    print(f"   ✅ Encontrados: {stats['encontrados']} ({stats['encontrados']*100//len(cpfs) if len(cpfs) > 0 else 0}%)")
    print(f"   ❌ Não encontrados: {stats['nao_encontrados']} ({stats['nao_encontrados']*100//len(cpfs) if len(cpfs) > 0 else 0}%)")
    print(f"   ⚠️  Erros: {stats['erros']}")
    print(f"   ⚠️  CPFs inválidos: {stats['cpfs_invalidos']}")
    print(f"\n💾 Arquivo salvo: {OUTPUT_ARQ}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()