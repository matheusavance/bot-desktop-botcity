
# Import externo Botcity
import os
import shutil
import sys
from twocaptcha import TwoCaptcha

# Import for the Desktop Bot
from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin
from botcity.plugins.files import BotFilesPlugin
from botcity.document_processing import *

# Import for integration with BotCity Maestro SDK
# from botcity.maestro import *

# Disable errors if we are not connected to Maestro
# BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    # maestro = BotMaestroSDK.from_sys_args()
    # Fetch the BotExecution with details from the task, including parameters
    # execution = maestro.get_execution()

    # Instância da classe DesktopBot
    bot = DesktopBot()

    # API Key TwoCaptcha
    api_key_twoCaptcha = "API_KEY"

    # URL site prefeitura
    url_site = 'https://iptu.prefeitura.sp.gov.br/'
    
    # Função JS aplicada no Chrome Developer
    funcao_js_img_captcha = 'base64 = document.getElementById("captchaImagem").src;function saveBase64AsFile(base64, fileName) {var link = document.createElement("a");link.setAttribute("href", base64);link.setAttribute("download", fileName);link.click();}saveBase64AsFile(base64, "download.jpg")'

    # Variável contadora de parcelas / variável do ano das parcelas que devem ser baixadas
    numero_parcela = int(input('\nDigite o número da parcela desejada: '))
    
    # Path dos boletos baixados
    path_boletos_renomeados = r"PATH_PASTA_BOLETOS"

    # Path img captcha
    path_arquivo_img_captcha = r"PATH_ARQUIVO_IMG"

    # Path boleto baixado
    path_arquivo_baixado = r"PATH_ARQUIVO_BAIXADO"

    # Path das planilhas usadas no processo
    path_planilha_base = r'PATH_PLANILHA_BASE'
    path_planilha_resultados = r'PATH_PLANILHA_RESULTADOS'

    # Transforma as planilhas de Status, Resultados e Base de dados em listas
    planilha_status = BotExcelPlugin('Status da extração').read(path_planilha_resultados)
    lista_planilha_status = planilha_status.as_list()

    planilha_resultados = BotExcelPlugin('Resultados da extração').read(path_planilha_resultados)
    lista_planilha_resultados = planilha_resultados.as_list()

    planilha_base = BotExcelPlugin("IPTU").read(path_planilha_base)
    lista_planilha_base = planilha_base.as_list()

    # Armazena o número da linha que deve ser preenchida nas planilhas
    linha_saida_status_extracao = len(lista_planilha_status) + 1
    linha_saida_resultado_extracao = len(lista_planilha_resultados) + 1

    #-------------#
    #   MÉTODOS   #
    #-------------#

    def login_site(url_site):
        # Entre no site da prefeitura
        bot.browse(url_site)
        bot.wait(3000)

    def verificacao_servico_captcha():
        bot.wait(1000)
        bot.control_a()
        bot.wait(1000)
        bot.control_c()     
        texto_pagina = bot.get_clipboard()
        
        if 'acesso ao serviço' in texto_pagina:
            sys.exit('> Erro no acesso ao serviço de CAPTCHA.\n')               

    def quebra_captcha_verificacao_bot(api_key_twoCaptcha, path_arquivo_img_captcha):
        bot.wait(1000)
        bot.control_a()
        bot.wait(1000)
        bot.control_c()     
        texto_pagina = bot.get_clipboard()
        
        if 'desafio' in texto_pagina:
            print('> Verificação de bot ocorrida\n')           
            solver = TwoCaptcha(api_key_twoCaptcha)
            try:
                bot.right_click_at(174, 176)
                bot.wait(500)
                bot.type_down()
                bot.wait(500)
                bot.type_down()
                bot.wait(500)
                bot.enter()
                bot.wait(4000)
                resultado_quebra_captcha = solver.normal(path_arquivo_img_captcha)
            
            except Exception as erroCaptcha:
                print(erroCaptcha)
                print('> Erro ao quebrar captcha. Quebrando novamente.\n')
            else:
                tolken_captcha = resultado_quebra_captcha['code']
                
                # Apaga o arquivo "kaptcha.png" armezenado na pasta de boletos baixados
                if os.path.isfile(path_arquivo_img_captcha):
                    os.remove(path_arquivo_img_captcha)
                
                # Clica no campo, preenche o tolken_captcha e clica em "submit"
                bot.control_w()
                bot.click_at(307,307)
                bot.type_key(tolken_captcha)
                bot.tab()
                bot.enter()
                bot.wait(1000)
                bot.control_w()

        else:
            print('> Verificação de bot NÃO ocorrida \n')          

    def quebra_captcha(api_key_twoCaptcha, path_arquivo_img_captcha):
        solver = TwoCaptcha(api_key_twoCaptcha)
        try:          
            bot.key_f12()
            bot.wait(2000)
            bot.type_key(funcao_js_img_captcha)
            bot.wait(2000)
            bot.enter()     
            bot.wait(2000)    
            bot.enter()
            bot.key_f12()
            bot.wait(4000)     
            resultado_quebra_captcha = solver.normal(path_arquivo_img_captcha)
        
        except Exception as erroCaptcha:
            print(erroCaptcha)
            print('> Erro ao quebrar captcha. Tentando quebra novamente\n')
            
            # Apaga o arquivo "kaptcha.png" armezenado na pasta de boletos baixados
            if os.path.isfile(path_arquivo_img_captcha):
                os.remove(path_arquivo_img_captcha)

            quebra_captcha(api_key_twoCaptcha, path_arquivo_img_captcha)
        else:
            tolken_captcha = resultado_quebra_captcha['code']
            
            # Apaga o arquivo "kaptcha.png" armezenado na pasta de boletos baixados
            if os.path.isfile(path_arquivo_img_captcha):
                os.remove(path_arquivo_img_captcha)

            print('> Captcha quebrado com sucesso.\n')

            return tolken_captcha
    
    def pesquisa_inscricao(inscricao_imobiliaria,numero_parcela):
        # Aperta Tab 3 vezes para encontrar o campo de "Número do Contribuinte"
        bot.tab()
        bot.tab()
        bot.tab()

        # Preenche o campo "Número do Contribuinte"
        for caracter in inscricao_imobiliaria:
            bot.kb_type(caracter)

        # Aperta Tab para passar para o campo "Parcela a ser impressa"
        bot.tab()

        # Enter para abrir o select do campo "Parcela a ser impressa"
        bot.enter()

        # Seleciona a parcela desejada
        for numero in range(1, numero_parcela):
            bot.type_down()

        # Aperta enter para selecionar a parcela encontrada
        bot.enter()

        # Aperta tab para encontrar o campo "Exercício(AAAA)"
        bot.tab()

        # Preenche o  campo "Exercício(AAAA)"
        bot.kb_type('2023')

        # Aperta tab para encontrar o campo "Código da imagem"
        bot.tab()

        # Resolve o text-captcha
        token_captcha = quebra_captcha(api_key_twoCaptcha, path_arquivo_img_captcha)

        # Preenche o campo "Código da imagem"
        for i in token_captcha: 
            bot.kb_type(i)           

        # Aperta tab duas vezes para encontrar o botão "GERAR 2ºVIA"               
        bot.tab()
        bot.tab()

        # Aperta enter para gerar 2ºvia
        bot.enter()
        bot.wait(1000)

    def verificacao_status_parcela(linha_saida_status_extracao, numero_parcela):
        bot.wait(1000)
        bot.control_a()
        bot.wait(1000)
        bot.control_c()     
        texto_pagina = bot.get_clipboard()
        
        if 'deve pagar' in texto_pagina:
            # Insere dados obtidos na planilha statusExtracao
            planilha_resultados.set_active_sheet('Status da extração')
            planilha_resultados.set_cell("A", linha_saida_status_extracao, id_apart, sheet="Status da extração")
            planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
            planilha_resultados.set_cell("C", linha_saida_status_extracao, cpf_cnpj, sheet="Status da extração")
            planilha_resultados.set_cell("D", linha_saida_status_extracao, numero_parcela, sheet="Status da extração")
            planilha_resultados.set_cell("E", linha_saida_status_extracao, "Resposta site: Nada deve pagar", sheet="Status da extração")
            planilha_resultados.write(path_planilha_resultados)

            print(f'> Parcela {numero_parcela}: Nada deve pagar"\n')
            print('------------------------------------------\n')

            # Fecha a guia
            bot.control_w()

            return True

        elif 'Número do contribuinte inválido' in texto_pagina:
            # Insere dados obtidos na planilha statusExtracao
            planilha_resultados.set_active_sheet('Status da extração')
            planilha_resultados.set_cell("A", linha_saida_status_extracao, id_apart, sheet="Status da extração")
            planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
            planilha_resultados.set_cell("C", linha_saida_status_extracao, cpf_cnpj, sheet="Status da extração")
            planilha_resultados.set_cell("D", linha_saida_status_extracao, numero_parcela, sheet="Status da extração")
            planilha_resultados.set_cell("E", linha_saida_status_extracao, "Resposta site: Número do contribuinte inválido", sheet="Status da extração")
            planilha_resultados.write(path_planilha_resultados)

            print(f'> Parcela {numero_parcela}: Número do contribuinte inválido"\n')
            print('------------------------------------------\n')

            # Fecha a guia
            bot.control_w()

            return True
        
        elif 'Código da imagem inválido' in texto_pagina:
            # Insere dados obtidos na planilha statusExtracao
            planilha_resultados.set_active_sheet('Status da extração')
            planilha_resultados.set_cell("A", linha_saida_status_extracao, id_apart, sheet="Status da extração")
            planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
            planilha_resultados.set_cell("C", linha_saida_status_extracao, cpf_cnpj, sheet="Status da extração")
            planilha_resultados.set_cell("D", linha_saida_status_extracao, numero_parcela, sheet="Status da extração")
            planilha_resultados.set_cell("E", linha_saida_status_extracao, "Resposta site: Código da imagem inválido", sheet="Status da extração")
            planilha_resultados.write(path_planilha_resultados)

            print(f'> Parcela {numero_parcela}: Código da imagem inválido"\n')
            print('------------------------------------------\n')

            # Fecha a guia
            bot.control_w()

            return True
        
        elif 'Não foram encontrados lançamentos' in texto_pagina:
            # Insere dados obtidos na planilha statusExtracao
            planilha_resultados.set_active_sheet('Status da extração')
            planilha_resultados.set_cell("A", linha_saida_status_extracao, id_apart, sheet="Status da extração")
            planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
            planilha_resultados.set_cell("C", linha_saida_status_extracao, cpf_cnpj, sheet="Status da extração")
            planilha_resultados.set_cell("D", linha_saida_status_extracao, numero_parcela, sheet="Status da extração")
            planilha_resultados.set_cell("E", linha_saida_status_extracao, "Resposta site: Não foram encontrados lançamentos", sheet="Status da extração")
            planilha_resultados.write(path_planilha_resultados)

            print(f'> Parcela {numero_parcela}: Não foram encontrados lançamentos"\n')
            print('------------------------------------------\n')

            # Fecha a guia
            bot.control_w()

            return True
        
        elif 'Confirmação' in texto_pagina:
            print(f'> Confirmação de emissão ocorrida após clicar em "Gerar 2ºvia\n')
            
            bot.tab()
            bot.tab()
            bot.enter()  

    def baixa_parcela():
        # Clica em "Imprimir"
        bot.tab()
        bot.wait(500)
        bot.enter()
        bot.wait(500)
        bot.enter()
        bot.wait(500)
        bot.enter()
    
    def verificacao_boleto_baixado(path_arquivo_baixado):
        if not os.path.exists(path_arquivo_baixado):
            print('> Erro do site ao tentar baixar boleto.\n')         
            print('------------------------------------------\n')
        
            # Insere dados obtidos na planilha statusExtracao
            planilha_resultados.set_active_sheet('Status da extração')
            planilha_resultados.set_cell("A", linha_saida_status_extracao, id_apart, sheet="Status da extração")
            planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
            planilha_resultados.set_cell("C", linha_saida_status_extracao, cpf_cnpj, sheet="Status da extração")
            planilha_resultados.set_cell("D", linha_saida_status_extracao, numero_parcela, sheet="Status da extração")
            planilha_resultados.set_cell("E", linha_saida_status_extracao, "Erro do site ao tentar baixar boleto.\n", sheet="Status da extração")
            planilha_resultados.write(path_planilha_resultados)

            # Fecha a guia
            bot.alt_f4()

            return True

    def renomea_boleto(inscricao_imobiliaria, numero_parcela, path_arquivo_baixado):
        # Verifica se há duplicidade de boleto. Se não existe, excluí o arquivo original, copia o boleto baixado e renomea o boleto copiado.
        path_arquivo_renomeado = r"C:\BotCity\bots\crisiptusaopaulotabas\Extração\Boletos renomeados"
        arquivo_renomeado = os.path.join(path_arquivo_renomeado,f"{inscricao_imobiliaria}_{numero_parcela}.pdf")
        
        if os.path.exists(arquivo_renomeado):
            os.remove(path_arquivo_baixado)
            print('> ATENÇÃO - Inscrição imobiliária duplicada na base.\n')
            bot.control_w()
            return True
        else:                         
            arquivo_copiado = shutil.copy2(path_arquivo_baixado, arquivo_renomeado)
            bot.wait(1500)
            os.remove(path_arquivo_baixado)
            print(f'> Parcela {numero_parcela}: Boleto baixado e renomeado\n')

            return arquivo_copiado
        
    def extrai_dados_boleto(path_arquivo_baixado):
        # Extrai dados do boleto
        reader = PDFReader()
        parser = reader.read_file(path_arquivo_baixado)
        _vencimento = parser.get_first_entry("VENCIMENTO")
        vencimento = parser.read(_vencimento, 1.51, 1.25, 2.29, 5.625)
        _autenticacao_mecanica = parser.get_first_entry("AUTENTICAÇÃO MECÂNICA")
        linha_digitavel = parser.read(_autenticacao_mecanica, -0.12, -6.75, 5.32, 5.625)
        linha_digitavel = linha_digitavel.replace(".","")
        linha_digitavel = linha_digitavel.replace(" ","")
        _parcela = parser.get_first_entry("PARCELA")
        parcela = parser.read(_parcela, 0.676056, 1, 3.028169, 5.875)
        _valor_a_pagar = parser.get_first_entry("VALOR A PAGAR")
        valor = parser.read(_valor_a_pagar, 0.843478, 1.375, 1.773913, 5.375)

        # Obtem código de barra (Lembrando que é necessário não possui espaço ou caracteres especiais)
        codigo_barras1 = linha_digitavel[:4]
        codigo_barras2 = linha_digitavel[32:]
        codigo_barras3 = linha_digitavel[4:9]+linha_digitavel[10:20]+linha_digitavel[21:31]
        codigo_barras = codigo_barras1 + codigo_barras2 + codigo_barras3

        return vencimento, parcela, valor, linha_digitavel, codigo_barras

    def exclui_imagens_desktop():
        path_desktop = r"C:\Users\niago\Desktop"
        for fname in os.listdir(path_desktop):
            if fname.endswith('.jpg') or fname.endswith('.png'):
                path_arquivo = os.path.join(path_desktop, fname)
                os.remove(path_arquivo)

    #------------------------#
    #   INÍCIO DO PROCESSO   #
    #------------------------#
    
    # Verifica se existem arquivos jpg ou png na área de trabalho. Caso exista, exclui-os
    exclui_imagens_desktop()

    # Inicia um loop que intera sob a planilha base para baixar os boletos
    for linha in lista_planilha_base:
        if "ID Apart" in str(linha[0]):
            continue
        else:
            # Armazena o valor da linha da plalinha
            id_apart = linha[0]
            inscricao_imobiliaria = linha[1]
            cpf_cnpj = linha[2]

            print("\n# DADOS ANALISADOS #\n")
            print("> ID Apart:", id_apart)
            print("> Inscrição imobiliária:", inscricao_imobiliaria)
            print("> CPF/CNPJ:", cpf_cnpj)
            print("> Linha de saída - Status:", linha_saida_status_extracao)
            print("> Linha de saída - Resultados:", linha_saida_resultado_extracao, '\n')
            print('------------------------------------------\n')

            # Entra no site da prefeitura
            login_site(url_site)

            # Caso ocorra um erro no serviço de captcha do site, o processo é parado
            verificacao_servico_captcha()

            # Caso haja uma verificação de bot, esta é resolvida
            quebra_captcha_verificacao_bot(api_key_twoCaptcha, path_arquivo_img_captcha)

            # Pesquisa pelo boleto
            pesquisa_inscricao(inscricao_imobiliaria, numero_parcela)    

            # Verifica a resposta do site após clicar em "Gerar 2ºvia'
            resposta_verificacao_status_parcela = verificacao_status_parcela(linha_saida_status_extracao, numero_parcela)
            if resposta_verificacao_status_parcela == True:            
                linha_saida_status_extracao += 1
                continue

            # Caso haja uma verificação de bot, esta é resolvida
            quebra_captcha_verificacao_bot(api_key_twoCaptcha, path_arquivo_img_captcha)

            # Baixa parcela
            baixa_parcela()

            # Verifica se o boleto foi baixado com sucesso
            resposta_verificacao_boleto_baixado = verificacao_boleto_baixado(path_arquivo_baixado)
            if resposta_verificacao_boleto_baixado == True:
                linha_saida_status_extracao += 1
                continue

            # Extrai dados do boleto
            vencimento, parcela, valor, linha_digitavel, codigo_barras = extrai_dados_boleto(path_arquivo_baixado)

            # Renomea o boleto baixado
            renomea_boleto(inscricao_imobiliaria, parcela, path_arquivo_baixado)

            # Insere dados obtidos na planilha statusExtracao
            planilha_resultados.set_active_sheet('Status da extração')
            planilha_resultados.set_cell("A", linha_saida_status_extracao, id_apart, sheet="Status da extração")
            planilha_resultados.set_cell("B", linha_saida_status_extracao, inscricao_imobiliaria, sheet="Status da extração")
            planilha_resultados.set_cell("C", linha_saida_status_extracao, cpf_cnpj, sheet="Status da extração")
            planilha_resultados.set_cell("D", linha_saida_status_extracao, numero_parcela, sheet="Status da extração")
            planilha_resultados.set_cell("E", linha_saida_status_extracao, "Boleto baixado", sheet="Status da extração")

            # Insere dados obtidos na planilha resultadoStatus
            planilha_resultados.set_active_sheet('Resultados da extração')
            planilha_resultados.set_cell("A", linha_saida_resultado_extracao, id_apart, sheet="Resultados da extração")
            planilha_resultados.set_cell("B", linha_saida_resultado_extracao, inscricao_imobiliaria, sheet="Resultados da extração")
            planilha_resultados.set_cell("C", linha_saida_resultado_extracao, cpf_cnpj, sheet="Resultados da extração")
            planilha_resultados.set_cell("D", linha_saida_resultado_extracao, numero_parcela, sheet="Resultados da extração")
            planilha_resultados.set_cell("E", linha_saida_resultado_extracao, vencimento, sheet="Resultados da extração")
            planilha_resultados.set_cell("F", linha_saida_resultado_extracao, valor, sheet="Resultados da extração")
            planilha_resultados.set_cell("G", linha_saida_resultado_extracao, linha_digitavel, sheet="Resultados da extração")
            planilha_resultados.set_cell("H", linha_saida_resultado_extracao, codigo_barras, sheet="Resultados da extração")
            planilha_resultados.write(path_planilha_resultados)

            # Incrementa as variáveis contadoras
            linha_saida_status_extracao += 1
            linha_saida_resultado_extracao += 1                        

            # Fecha a guia 
            bot.control_w()

            print('\n------------------------------------------\n') 
   
if __name__ == '__main__':
    main()