from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
import pandas as pd
from datetime import datetime
import os
import bcrypt
import base64
import re # LINHA ADD NA MÃO
import folium          #  LINHA PARA O MAPA
from folium import Choropleth       #  LINHA PARA O MAPA
import json       #  LINHA PARA O MAPA


app = Flask(__name__)
app.secret_key = "chave_secreta"  # Para mensagens flash
#DATA_DIR = r"C:\Users\Fabio\Documents\Estudos\Python\SitePesquisa\dados_excel"  # Ajuste o caminho
DATA_DIR = r"C:\Users\Fabio\Documents\Estudos\Python\P&R\dados_excel"

# CODIGO PARA UPLOAD
UPLOAD_FOLDER = r"C:\Users\Fabio\Documents\Estudos\Python\P&R\static\uploads"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
MAX_FILE_SIZE = 2 * 1024 * 1024  # 2 MB em bytes

# Função para carregar uma planilha
def load_excel(file_name):
    try:
        return pd.read_excel(os.path.join(DATA_DIR, file_name), engine='openpyxl')
    except FileNotFoundError:
        return pd.DataFrame()

# Função para salvar uma planilha
def save_excel(df, file_name):
    df.to_excel(os.path.join(DATA_DIR, file_name), index=False, engine='openpyxl')

# Rota para página inicial (login)
# LOGIN PARA TESTE OFFLINE. TIREI PARA O LOGIN ABAIXO ONLINE
# @app.route("/", methods=["GET", "POST"])
#def index():
    #if request.method == "POST":
        #nome_usuario = request.form["nome_usuario"]
        #senha = request.form["senha"].encode('utf-8')
        #usuarios = load_excel("Usuarios.xlsx")
        #usuario = usuarios[usuarios["Nome_Usuario"] == nome_usuario]
        #print(f"Usuário encontrado: {usuario.to_dict() if not usuario.empty else 'Nenhum'}")  # Depuração
        #if not usuario.empty:
            #senha_hashed = usuario.iloc[0]["Senha"].encode('utf-8')
            #print(f"Senha no banco: {senha_hashed}")  # Depuração
            #if bcrypt.checkpw(senha, senha_hashed):
                #session["nome_usuario"] = nome_usuario  # Armazena o usuário na sessão
                #flash("Login bem-sucedido!", "success")
                #tipo = usuario.iloc[0]["Tipo"]
                #if tipo == "Admin":
                    #return redirect(url_for("questionario"))
                    ##return redirect(url_for("admin"))    # LINHA ORIGINAL
                #elif tipo == "Consultor":
                    #return redirect(url_for("questionario"))
                #else:
                    #return redirect(url_for("cadastro_aluno"))
            #else:
                #flash("Senha inválida!", "danger")
        #else:
            #flash("Usuário não encontrado!", "danger")
    #return render_template("index.html")

# NOVA ROTA DE LOGIN PARA TESTE ONLINE
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nome_usuario = request.form["nome_usuario"].strip()
        senha = request.form["senha"].encode('utf-8')

        # Usuários de teste (remova depois ou use planilha criptografada)
        usuarios_teste = {
            "admin": bcrypt.hashpw("admin10".encode('utf-8'), bcrypt.gensalt()),
            "teste": bcrypt.hashpw("1234".encode('utf-8'), bcrypt.gensalt())
        }

        if nome_usuario in usuarios_teste:
            senha_hashed = usuarios_teste[nome_usuario]
            if bcrypt.checkpw(senha, senha_hashed):
                session["nome_usuario"] = nome_usuario
                flash("Login bem-sucedido!", "success")
                return redirect(url_for("questionario"))  # ou "admin"
            else:
                flash("Senha inválida!", "danger")
        else:
            flash("Usuário não encontrado!", "danger")

    return render_template("index.html")


# Rota para o questionário do fiscal

@app.route("/questionario", methods=["GET", "POST"])
def questionario():
    if "nome_usuario" not in session:
        flash("Faça login para acessar o questionário!", "danger")
        return redirect(url_for("index"))

    # nucleos = load_excel("Nucleos.xlsx")   # LINHA ORIGINAL TABELA NUCLEOS

    # Carrega e PADRONIZA os nomes das colunas dos bairros
    bairros_df = load_excel("TodosBairros.xlsx")
    bairros = bairros_df.rename(columns={
        "Bairro": "Nome",
        "Latitude": "Lat",
        "Longitude": "Lon",
        "Reg": "Reg"
    })

    ##bairros = load_excel("TodosBairros.xlsx")
    alunos = load_excel("Alunos.xlsx")
    professores = load_excel("Professores.xlsx")
    projetos = load_excel("Projetos.xlsx")  # PLANILHA PROJETOS
    turnos = load_excel("Turnos.xlsx")  # PLANILHA TURNOS
    
    secoes_df = load_excel("FortalZonSecLoc.xlsx")
    
    # Remove espaços em branco nos nomes das colunas (problema comum no Excel)
    secoes_df.columns = secoes_df.columns.str.strip()
    
    # Lista única de zonas ordenadas
    zonas_unicas = sorted(secoes_df["Zona"].unique())
    
    # DataFrame completo pra usar nas buscas
    secoes = secoes_df
    
    print("Zonas únicas carregadas:", zonas_unicas)
    print("Total de seções:", len(secoes))
    ## FIN CARREGA ZONAS E SECOES 

    #print(f"Núcleos carregados: {nucleos.to_dict('records')}")  # Debug
    print("Passo 1 ")
    
    if request.method == "POST":
        print(f"Dados recebidos do formulário: {request.form.to_dict(flat=False)}")  # Debug completo
        print(f"Chaves do formulário: {list(request.form.keys())}")  # Debug chaves
        print("Passo 2 ")

        # PEGA TODAS AS VARIÁVEIS PRIMEIRO
        data = request.form.get("data")
        horario_fim = request.form.get("horario_fim")
    
        # AGORA FAZ O DEBUG
        print(f"VALOR BRUTO DE 'data' que chegou do form: '{data}' (tipo: {type(data)})")
        print(f"VALOR BRUTO DE 'horario_fim' que chegou do form: '{horario_fim}' (tipo: {type(horario_fim)})")

        data = request.form.get("data")
        ## projeto_codigo = request.form.get("projeto")  # PLANILHA PROJETOS
        ## nucleo_codigo = request.form.get("nucleo")    ## LINHA ORIGINAL
        biometria = request.form.get("biometria", "Não")
        #nascimento = request.form.get("nascimento") # CODIGO ANTIGO; PEGA A DATA DE NASCIMENTO
        #cpf = request.form.get("cpf")
        cpf = request.form.get("cpf", "").strip()
        bairro_codigo = request.form.get("bairro")
        pai = request.form.get("pai")
        mae = request.form.get("mae")
        #horario_inicio = request.form.get("horario_inicio")    # COD ANTIGO; PEGA O HORARIO DE INICIO
        horario_fim = request.form.get("horario_fim")         # COD ANTIGO; PEGA O HORARIO DE FIM
        #if secao_selecionada:
            ## Busca na planilha de seções
            #secao_info = secoes[secoes["Secao"] == secao_selecionada]
            #if not secao_info.empty:
                #zona_selecionada = secao_info["Zona"].iloc[0]
            #else:
                #zona_selecionada = ""  # Caso não encontre (improvável, mas seguro)
        
        ## NOVO CÓDIGO PARA PEGAR ZONA E DEPOIS SEÇÃO
        zona_selecionada = request.form.get("zona")  # agora vem direto do select
        secao_selecionada = request.form.get("secao")
        
        # Opcional: valida se a seção pertence à zona (segurança extra)
        if secao_selecionada and zona_selecionada:
            try:
                secao_int = int(secao_selecionada)
                zona_int = int(zona_selecionada)
                if not secoes[(secoes["Secao"] == secao_int) & (secoes["Zona"] == zona_int)].empty:
                    # Tudo certo
                    pass
                else:
                    flash("Seção não pertence à zona selecionada!", "danger")
                    # return render...
            except:
                pass

        # CODIGO NOVO ONDE PADRONIZA CAMPO NOME
        #titulo = request.form.get("titulo", "")[:16]   # LINHA ANTIGA, SÓ ACEITA 16 DIGITOS
        titulo = request.form.get("titulo", "").strip()

        # === PADRONIZA NOMES ===
        nome = request.form.get("nome", "").strip()
        if nome:
            nome = nome.title()[:100]

        pai = request.form.get("pai", "").strip()
        if pai:
            pai = pai.title()[:100]

        mae = request.form.get("mae", "").strip()
        if mae:
            mae = mae.title()[:100]
        
        endereco = request.form.get("endereco", "").strip()
        if endereco:
            endereco = endereco.title()[:100]
        # FIM CODIGO NOVO ONDE PADRONIZA CAMPO NOME

        # === INICIO VALIDAÇÃO DE IDADE NO NASCIMENTO ===
        nascimento = request.form.get("nascimento")
        if nascimento:
            try:
                data_nasc = datetime.strptime(nascimento, "%Y-%m-%d")
                hoje = datetime.now()
                
                idade = hoje.year - data_nasc.year - ((hoje.month, hoje.day) < (data_nasc.month, data_nasc.day))
                
                if idade < 16:
                    flash("Data de nascimento inválida: a pessoa deve ter pelo menos 16 anos.", "danger")
                    return render_template("questionario.html", 
                                          alunos=alunos.to_dict('records'), 
                                          bairros=bairros.to_dict('records'), 
                                          professores=professores.to_dict('records'),
                                          projetos=projetos.to_dict('records'),
                                          secoes=secoes.to_dict('records') if not secoes.empty else [],
                                          zonas=zonas_unicas,
                                          turnos=turnos.to_dict('records'),
                                          usuario=session.get("nome_usuario", ""))
                elif idade > 100:
                    flash("Data de nascimento inválida: a pessoa não pode ter mais de 100 anos.", "danger")
                    return render_template("questionario.html", 
                                          alunos=alunos.to_dict('records'), 
                                          bairros=bairros.to_dict('records'), 
                                          professores=professores.to_dict('records'),
                                          projetos=projetos.to_dict('records'),
                                          secoes=secoes.to_dict('records') if not secoes.empty else [],
                                          zonas=zonas_unicas,
                                          turnos=turnos.to_dict('records'),
                                          usuario=session.get("nome_usuario", ""))
            except ValueError:
                flash("Data de nascimento inválida! Use o formato correto.", "danger")
                return render_template("questionario.html", 
                                      alunos=alunos.to_dict('records'), 
                                      bairros=bairros.to_dict('records'), 
                                      professores=professores.to_dict('records'),
                                      projetos=projetos.to_dict('records'),
                                      secoes=secoes.to_dict('records') if not secoes.empty else [],
                                      zonas=zonas_unicas,
                                      turnos=turnos.to_dict('records'),
                                      usuario=session.get("nome_usuario", ""))
        # === FIN VALIDAÇÃO DE IDADE NO NASCIMENTO ===

        # Verificar campos obrigatórios
        required_fields = {
            #"data": data,
            ##"projeto": projeto_codigo, ## CAMPO EXCLUIDO
            ##"nucleo": nucleo_codigo,
            "bairro": bairro_codigo,
            ##"prof_uniforme": prof_uniforme
        }
        for field_name, field_value in required_fields.items():
            if not field_value:
                print(f"Campo obrigatório ausente: {field_name}")  # Debug
                flash(f"Campo {field_name} é obrigatório!", "danger")
                print("Passo 3 ")
                return render_template("questionario.html", 
                                    alunos=alunos.to_dict('records'), 
                                    ##nucleos=nucleos.to_dict('records'), 
                                    bairros=bairros.to_dict('records'), 
                                    professores=professores.to_dict('records'),
                                    projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                    secoes=secoes.to_dict('records'),   # CAMPOS DA PLANILHA C/SECOES ZONAS LOCAIS
                                    zonas=zonas_unicas,                # <--- NOVO CODIGO PARA CONSULTAR POR ZONA
                                    turnos=turnos.to_dict('records'),
                                    usuario=session.get("nome_usuario", ""))
        
        # Validar quantidade_alunos
        ## try:
            ## quantidade_alunos = int(quantidade_alunos)
            ## if not (0 <= quantidade_alunos <= 99999):
                ## print(f"Meta de Atendimento Fora do Intervalo: {quantidade_alunos}")  # Debug
                ## flash("Meta de Atendimento Deve Estar Entre 0 e 99999!", "danger")
                ## return render_template("questionario.html", 
                                    ## alunos=alunos.to_dict('records'), 
                                    ## ##nucleos=nucleos.to_dict('records'), 
                                    ## bairros=bairros.to_dict('records'), 
                                    ## professores=professores.to_dict('records'),
                                    ## projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                    ## turnos=turnos.to_dict('records'),
                                    ## usuario=session.get("nome_usuario", ""))
        ## except ValueError:
            ## print(f"Meta Inválida: {quantidade_alunos}")  # Debug
            ## flash("Meta Deve Ser um Número Válido!", "danger")
            ## return render_template("questionario.html", 
                                ## alunos=alunos.to_dict('records'), 
                                ## ##nucleos=nucleos.to_dict('records'), 
                                ## bairros=bairros.to_dict('records'), 
                                ## professores=professores.to_dict('records'),
                                ## projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                ## turnos=turnos.to_dict('records'),
                                ## usuario=session.get("nome_usuario", ""))

        # Validar data - CODIGO ANTIGO 
        #try:
            #data_obj = datetime.strptime(data, "%Y-%m-%d")
            #data_formatada = data_obj.strftime("%d/%m/%Y")
        #except ValueError as e:
            #print(f"Erro na validação de data: {str(e)}")  # Debug
            #flash("Data inválida! Use o formato AAAA-MM-DD.", "danger")
            #return render_template("questionario.html", 
                                #alunos=alunos.to_dict('records'), 
                                ###nucleos=nucleos.to_dict('records'), 
                                #bairros=bairros.to_dict('records'), 
                                #professores=professores.to_dict('records'),
                                #projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                #turnos=turnos.to_dict('records'),
                                #usuario=session.get("nome_usuario", ""))
        
        # Validar data no formato brasileiro dd/mm/yyyy   -   CODIGO NOVO
        try:
            data_obj = datetime.strptime(data, "%d/%m/%Y")
            data_formatada = data_obj.strftime("%d/%m/%Y")  # Salva como dd/mm/yyyy na planilha
        except ValueError:
            try:
                # Tenta o formato americano como fallback
                data_obj = datetime.strptime(data, "%Y-%m-%d")
                data_formatada = data_obj.strftime("%d/%m/%Y")
            except ValueError:
                flash("Data inválida! Use o formato DD/MM/AAAA.", "danger")
                return render_template("questionario.html", 
                                    alunos=alunos.to_dict('records'), 
                                    bairros=bairros.to_dict('records'), 
                                    professores=professores.to_dict('records'),
                                    projetos=projetos.to_dict('records'),
                                    secoes=secoes.to_dict('records'),   # CAMPOS DA PLANILHA C/SECOES ZONAS LOCAIS
                                    zonas=zonas_unicas,                # <--- NOVO CODIGO PARA CONSULTAR POR ZONA
                                    turnos=turnos.to_dict('records'),
                                    usuario=session.get("nome_usuario", ""))

        # Validar formato do horário
        time_pattern = r"^[0-2][0-9]:[0-5][0-9](:[0-5][0-9])?$"
        # if not re.match(time_pattern, horario_inicio) or not re.match(time_pattern, horario_fim) or not re.match(time_pattern, hora1OK) or not re.match(time_pattern, hora2OK): # LINHA ORIGINAL
        ##if not re.match(time_pattern, horario_inicio) or not re.match(time_pattern, horario_fim):
        if not re.match(time_pattern, horario_fim) or not re.match(time_pattern, horario_fim):
            print(f"Horário inválido - Início: {horario_fim}")  # Debug
            #print(f"Horário inválido - Início: {horario_fim}, Fim: {horario_fim}, Hora1: {hora1OK}, Hora2: {hora2OK}")  # Debug
            flash("Horários inválidos! Use o formato HH:MM.", "danger")
            return render_template("questionario.html", 
                                alunos=alunos.to_dict('records'), 
                                ##nucleos=nucleos.to_dict('records'), 
                                bairros=bairros.to_dict('records'), 
                                professores=professores.to_dict('records'),
                                projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                secoes=secoes.to_dict('records'),   # CAMPOS DA PLANILHA C/SECOES ZONAS LOCAIS
                                zonas=zonas_unicas,                # <--- NOVO CODIGO PARA CONSULTAR POR ZONA
                                turnos=turnos.to_dict('records'),
                                usuario=session.get("nome_usuario", ""))
        
        # Validar aluno, bairro e professor
        try:
            #aluno_codigo = int(aluno_codigo) # LINHA ORIGINAL
            ##projeto_codigo = int(projeto_codigo)  # PLANILHA PROJETO
            ##turno_codigo = int(turno_codigo)  # PLANILHA TURNO
            ##professor_codigo = int(professor_codigo)
            #estagiario_codigo = int(estagiario_codigo)
            ##nucleo_codigo = int(nucleo_codigo)
            bairro_codigo = int(bairro_codigo)
            bairro = bairros[bairros["Codigo"] == bairro_codigo]
            if bairro.empty:
                raise IndexError("Bairro Não Encontrado")
            bairro_nome = bairro["Nome"].iloc[0]
            bairro_lat = bairro["Lat"].iloc[0]
            bairro_lon = bairro["Lon"].iloc[0]
            bairro_reg = bairro["Reg"].iloc[0] # LINHA CRIADA PRA SALVAR REGIONAL 3-12-25
            print(f"Bairro Selecionado: Código={bairro_codigo}, Nome={bairro_nome}, Lat={bairro_lat}, Lon={bairro_lon}")
        except (ValueError, IndexError, KeyError) as e:
            print(f"Erro na validação: {str(e)}")  # Debug
            flash("Projeto, Núcleo ou Professor Inválido!", "danger")
            return render_template("questionario.html", 
                                alunos=alunos.to_dict('records'), 
                                ##nucleos=nucleos.to_dict('records'), 
                                bairros=bairros.to_dict('records'), 
                                professores=professores.to_dict('records'),
                                projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                secoes=secoes.to_dict('records'),   # CAMPOS DA PLANILHA C/SECOES ZONAS LOCAIS
                                zonas=zonas_unicas,                # <--- NOVO CODIGO PARA CONSULTAR POR ZONA
                                turnos=turnos.to_dict('records'),
                                usuario=session.get("nome_usuario", ""))
        
        ## SEM USO
        # Validar tamanho do arquivo de foto
        ##foto_path = ""
        ##if foto and foto.filename:
            ##foto.seek(0, os.SEEK_END)
            ##file_size = foto.tell()
            ##foto.seek(0)
            ##if file_size > MAX_FILE_SIZE:
                ##flash("Arquivo muito grande! O tamanho máximo é 2 MB.", "danger")
                ##return render_template("questionario.html", 
                                    ##alunos=alunos.to_dict('records'), 
                                    ####nucleos=nucleos.to_dict('records'), 
                                    ##bairros=bairros.to_dict('records'), 
                                    ##professores=professores.to_dict('records'),
                                    ##projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                    ##turnos=turnos.to_dict('records'),
                                    ##usuario=session.get("nome_usuario", ""))
            ##filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{foto.filename}"
            ##filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            ##try:
                ##foto.save(filepath)
                ##foto_path = f"static/uploads/{filename}"
            ##except Exception as e:
                ##flash(f"Erro ao salvar a foto: {str(e)}", "danger")
                ##return render_template("questionario.html", 
                                    ##alunos=alunos.to_dict('records'), 
                                    ####nucleos=nucleos.to_dict('records'), 
                                    ##bairros=bairros.to_dict('records'), 
                                    ##professores=professores.to_dict('records'),
                                    ##projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                                    ##turnos=turnos.to_dict('records'),
                                    ##usuario=session.get("nome_usuario", ""))
        
        # Salvar questionário
        questionarios = load_excel("Questionarios.xlsx")

        # === INICIO FORÇA COLUNAS COMO TEXTO (evita notação científica) ===
        if not questionarios.empty:
            if "NumTit" in questionarios.columns:
                questionarios["NumTit"] = questionarios["NumTit"].astype(str)
            if "CPF" in questionarios.columns:
                questionarios["CPF"] = questionarios["CPF"].astype(str)
        # === FIN FORÇA COLUNAS COMO TEXTO (evita notação científica) ===

        new_questionario = pd.DataFrame([{
            "Codigo": (questionarios["Codigo"].max() + 1) if not questionarios.empty else 1,
            "Data": data_formatada,
            "Nome": nome,
            #"NumTit": titulo,      #  SALVA CAMPO NUM TITULO SEM TRATAMENTO. PODE DAR ERRO QUANDO FOR MAIOR Q 15 DIGITOS
            #"NumTit": str(titulo).zfill(16),          # <--- CONVERTE PRA TEXTO e completa com zeros à esquerda (opcional)
            "NumTit": str(titulo),          # <--- CONVERTE PRA TEXTO
            "Biometria": 1 if biometria == "Sim" else 0,
            "DtNasc": nascimento,
            #"CPF": cpf,          #  SALVA CAMPO CPF SEM TRATAMENTO. PODE DAR ERRO QUANDO FOR MAIOR Q 15 DIGITOS
            "CPF": str(cpf),         # <--- CONVERTE PRA TEXTO
            "Zona": zona_selecionada or "",
            "Secao": secao_selecionada or "",
            "Pai": pai,
            "Mãe": mae,
            ##"Nome_do_Projeto": projeto_nome,   # PLANILHA PROJETO
            ##"Nome_do_Nucleo": nucleo_nome,
            "Endereço": endereco,
            "Bairro": bairro_nome,
            ##"Endereço_do_Nucleo": nucleo_endereco,
            "Latitude": bairro_lat,
            ##"Bairro_do_Nucleo": nucleo_bairro,
            "Longitude": bairro_lon,
            "Regional": bairro_reg,
            #"Nome_do_Aluno": aluno_nome,
            #"HorarioInicioAula": horario_inicio,
            ##"HorarioAtividade": horario_inicio,
            ##"Turno": turno_nome,   # PLANILHA TURNO
            "HorarioVisita": horario_fim,
            #"Voto": fiscal,
            #"Observacao": observacao,
            ##"Modalidade": ",".join(modalidades) if modalidades else "", # LISTA MODALIDADES
            ##"ProfPresente": 1 if prof_presente == "Sim" else 0,
            ## "Consideracoes": txt_fim,
            #"Material": ",".join(materiais) if materiais else "", # LINHA DA ANTIGA RELACAO MATL ESPORTIVOS
            # "Observacao": observacao,
            #"Foto": foto_path,
            "Usuario": session["nome_usuario"],
            ##"Foto": foto_path
        }])
        print(f"Salvando Questionário: {new_questionario.to_dict('records')}")  # Debug
        questionarios = pd.concat([questionarios, new_questionario], ignore_index=True)
        # Garante que as colunas sejam string ANTES de salvar
        questionarios["NumTit"] = questionarios["NumTit"].astype(str)
        questionarios["CPF"] = questionarios["CPF"].astype(str)
        save_excel(questionarios, "Questionarios.xlsx")
        
        flash("Questionário Salvo Com Sucesso!", "success")
        return redirect(url_for("questionario"))
    
    return render_template("questionario.html", 
                          alunos=alunos.to_dict('records'), 
                          ##nucleos=nucleos.to_dict('records'), 
                          bairros=bairros.to_dict('records'), 
                          professores=professores.to_dict('records'),
                          projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                          secoes=secoes.to_dict('records'),   # CAMPOS DA PLANILHA C/SECOES ZONAS LOCAIS
                          zonas=zonas_unicas,                # <--- NOVO CODIGO PARA CONSULTAR POR ZONA
                          turnos=turnos.to_dict('records'),
                          usuario=session.get("nome_usuario", ""))

# ROTA PARA ADMIN

@app.route("/admin", methods=["GET", "POST"])
def admin():
    if "nome_usuario" not in session:
        flash("Faça login para acessar!", "danger")
        return redirect(url_for("index"))
    
    # Verifica se é Admin
    usuarios_df = load_excel("Usuarios.xlsx")
    current_user = usuarios_df[usuarios_df["Nome_Usuario"] == session["nome_usuario"]]
    if current_user.empty or current_user.iloc[0]["Tipo"] != "Admin":
        flash("Acesso negado: apenas Admin pode gerenciar usuários.", "danger")
        return redirect(url_for("questionario"))
    
    mensagem = None
    
    if request.method == "POST":
        acao = request.form.get("acao")
        codigo = request.form.get("codigo")
        
        if acao == "adicionar" or acao == "editar":
            nome_usuario = request.form.get("nome_usuario").strip()
            nome_completo = request.form.get("usuario").strip().title()
            tipo = request.form.get("tipo")
            dica = request.form.get("dica").strip()
            senha_plana = request.form.get("senha")
            
            # Validações
            if not nome_usuario or not nome_completo or not tipo or not senha_plana:
                flash("Todos os campos são obrigatórios!", "danger")
            elif usuarios_df[usuarios_df["Nome_Usuario"] == nome_usuario].any().any() and acao == "adicionar":
                flash("Nome de usuário já existe!", "danger")
            else:
                # Criptografa a senha
                senha_hashed = bcrypt.hashpw(senha_plana.encode('utf-8'), bcrypt.gensalt())
                
                if acao == "adicionar":
                    novo_codigo = (usuarios_df["Codigo"].max() + 1) if not usuarios_df.empty else 1
                    nova_linha = pd.DataFrame([{
                        "Codigo": novo_codigo,
                        "Nome_Usuario": nome_usuario,
                        "Usuario": nome_completo,
                        "Tipo": tipo,
                        "Dica": dica,
                        "Senha": senha_hashed.decode('utf-8'),
                        "Data_cadastro": datetime.now().strftime("%d/%m/%Y")
                    }])
                    usuarios_df = pd.concat([usuarios_df, nova_linha], ignore_index=True)
                    flash("Usuário adicionado com sucesso!", "success")
                
                elif acao == "editar":
                    codigo_int = int(codigo)
                    usuarios_df.loc[usuarios_df["Codigo"] == codigo_int, "Nome_Usuario"] = nome_usuario
                    usuarios_df.loc[usuarios_df["Codigo"] == codigo_int, "Usuario"] = nome_completo
                    usuarios_df.loc[usuarios_df["Codigo"] == codigo_int, "Tipo"] = tipo
                    usuarios_df.loc[usuarios_df["Codigo"] == codigo_int, "Dica"] = dica
                    if senha_plana:  # só altera senha se preenchida
                        usuarios_df.loc[usuarios_df["Codigo"] == codigo_int, "Senha"] = senha_hashed.decode('utf-8')
                    flash("Usuário atualizado com sucesso!", "success")
                
                save_excel(usuarios_df, "Usuarios.xlsx")
        
        elif acao == "excluir":
            codigo_int = int(codigo)
            usuarios_df = usuarios_df[usuarios_df["Codigo"] != codigo_int]
            # Reindexa códigos
            usuarios_df["Codigo"] = range(1, len(usuarios_df) + 1)
            save_excel(usuarios_df, "Usuarios.xlsx")
            flash("Usuário excluído com sucesso!", "success")
    
    # Carrega usuários atualizados
    usuarios_df = load_excel("Usuarios.xlsx")
    usuarios = usuarios_df.to_dict('records')
    
    return render_template("admin.html", usuarios=usuarios, usuario_logado=session["nome_usuario"])

#import folium
#from folium import Choropleth
#import json

@app.route("/mapa")
def mapa():
    if "nome_usuario" not in session:
        flash("Faça login para acessar!", "danger")
        return redirect(url_for("index"))
    
    # Verifica se é Admin (ou quem você quiser que veja o mapa)
    usuarios_df = load_excel("Usuarios.xlsx")
    current_user = usuarios_df[usuarios_df["Nome_Usuario"] == session["nome_usuario"]]
    
    # TESTA SE USUARIO É VAZIO
    if current_user.empty:
        flash("Usuário não encontrado.", "danger")
        return redirect(url_for("index"))

    # TESTA O TIPO DE USUARIO PARA ACESSAR HTML MAPA
    #if current_user.empty or current_user.iloc[0]["Tipo"] != "Admin":
        #flash("Acesso negado: apenas Admin pode ver o mapa.", "danger")
        #return redirect(url_for("questionario"))
    
    tipo_usuario = current_user.iloc[0]["Tipo"] if not current_user.empty else ""
    if tipo_usuario not in ["Admin", "Consultor"]:
        flash("Acesso negado: apenas Admin e Consultor podem ver o mapa.", "danger")
        return redirect(url_for("questionario"))

    # Carrega os dados
    questionarios = load_excel("Questionarios.xlsx")
    if questionarios.empty:
        flash("Nenhum cadastro ainda para mostrar no mapa.", "info")
        return render_template("mapa_vazio.html")  # ou só uma mensagem
    
    # Conta votos por bairro
    votos_por_bairro = questionarios["Bairro"].value_counts().to_dict()
    
    # Carrega o GeoJSON dos bairros de Fortaleza
    geojson_path = os.path.join(os.path.dirname(__file__), "Fortal-Bai.geojson")
    with open(geojson_path, "r", encoding="utf-8") as f:
        bairros_geo = json.load(f)
    
    # Cria mapa centrado em Fortaleza
    m = folium.Map(location=[-3.77777, -38.5434], zoom_start=12, tiles="OpenStreetMap")
    #mapa = folium.Map(location=[-3.77777, -38.5434], tiles='cartodbpositron', zoom_start=12)
    
    # Calcula o total de votos; cada linha é 1 voto
    total_votos = len(questionarios)

    # Adiciona o choropleth (mapa de calor por bairro) - ÚNICA CAMADA
    choropleth = Choropleth(
        geo_data=bairros_geo,
        data=pd.Series(votos_por_bairro),
        key_on="feature.properties.Nome",  # ajuste se necessário (ex: NM_BAIRRO)
        #fill_color="PuRd",                 # roxo claro → escuro (ou Blues, Greens, etc.)
        fill_color="YlOrRd",              # amarelo → laranja → vermelho (escuro)
        fill_opacity=0.4,
        line_opacity=0.9,
        nan_opacity=0.1,
        #legend_name="Quantidade de Votos por Bairro",
        #legend_name=f"Votos por Bairro (Total: {total_votos:,} votos)".replace(",", "."),
        name="Votos por Bairro",
        highlight=True,
        smooth_factor=0,
        #show=False,  # <--- DESATIVA A LEGENDA AUTOMÁTICA
    ).add_to(m)

    # Adiciona tooltip/popup automático do Choropleth (nome do bairro + votos)
    for feature in bairros_geo["features"]:
        bairro_nome = feature["properties"]["Nome"]  # ajuste se necessário
        votos = votos_por_bairro.get(bairro_nome, 0)
        feature["properties"]["votos"] = f" {votos}"
        #feature["properties"]["votos"] = f"Votos: {votos}"

    # Configura o geojson do Choropleth pra mostrar popup
    # CODIGO ANTIGO COM FONTE PEQUENA DO TOOLTIP
    #choropleth.geojson.add_child(
        #folium.features.GeoJsonTooltip(
            #fields=["Nome", "votos"],
            #aliases=["Bairro:", "Votos:"],
            #localize=True,
            #sticky=False,
            #labels=True,
            #style="""
                #background-color: white;
                #border: 2px solid black;
                #border-radius: 3px;
                #box-shadow: 3px;
            #"""
        #)
    #)
    # FIM

    # INI Configura o geojson do Choropleth pra mostrar popup
    # CODIGO NOVO
    choropleth.geojson.add_child(
        folium.features.GeoJsonTooltip(
            fields=["Nome", "votos"],
            aliases=["Bairro:", "Votos:"],
            localize=True,
            sticky=False,
            labels=True,
            style="""
                background-color: white;
                border: 2px solid #333;
                border-radius: 8px;
                box-shadow: 3px 3px 10px rgba(0,0,0,0.3);
                padding: 10px;
                font-family: Arial, sans-serif;
                font-size: 16px;          /* <--- FONTE MAIOR (de 12-14px padrão pra 16px) */
                font-weight: bold;
                color: #333;
            """,
            max_width=800
        )
    )

     # === LEGENDA MANUAL NO CANTO DIREITO COM ICONES COLORIDOS===
    #legend_html = f'''
    #<div style="
        #position: fixed; 
        #top: 110px; right: 10px; width: 100px;             /* <--- TESTE DE DIMENSOES */
        #background-color: white;
        #border:2px solid grey; 
        #border-radius:8px;
        #padding: 10px;
        #font-size:14px; 
        #z-index:1000;
        #box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    #">
        #<div style="font-weight:bold; text-align:center; margin-bottom:8px;">
            #Votos por Bairro<br>
            #<small>Total: {total_votos:,} </small>
        #</div>
        #<i style="background:#ffffb2; width:18px; height:18px; float:left; margin-right:8px; opacity:0.7;"></i><small>Poucos</small><br>
        #<i style="background:#fecc5c; width:18px; height:18px; float:left; margin-right:8px; opacity:0.7;"></i><small></small><br>
        #<i style="background:#fd8d3c; width:18px; height:18px; float:left; margin-right:8px; opacity:0.7;"></i><small>Médios</small><br>
        #<i style="background:#e31a1c; width:18px; height:18px; float:left; margin-right:8px; opacity:0.7;"></i><small>Muitos</small>
    #</div>
    #'''

    # === LEGENDA MANUAL NO CANTO DIREITO COM ICONES COLORIDOS===
    legend_html = f'''
    <div style="
        position: fixed; 
        top: 36px; right: 5px; width: 190px;             /* <--- TESTE DE DIMENSOES */
        <!--background-color: white;-->
        background-color: rgba(255, 255, 255, 0.7);   /* branco com 70% de opacidade */
        border:2px solid grey; 
        border-radius:8px;
        padding: 4px;
        font-size:16px; 
        z-index:1000;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    ">
        <div style="font-weight:bold; text-align:center; margin-bottom:8px;">
            <!--Votos por Bairro<br>-->
            <small>Total de Votos: {total_votos:,} </small>
        </div>
    </div>
    '''

    m.get_root().html.add_child(folium.Element(legend_html))
    # ================================================

    # Adiciona controle de camadas (só vai ter "Votos por Bairro")
    folium.LayerControl().add_to(m)

    # === LOGO ===
    logo_nome = 'logo.png'
    with open(logo_nome, "rb") as imagem_file:
        encoded_string = base64.b64encode(imagem_file.read()).decode()

    logo_html = f'''
    <div style="
        position: fixed;
        top: 635px;                            /* <--- TESTE ALTURA DA LOGO; ERA 15 */
        left: 10px;
        width: 60px;                          /* <--- TESTE TAMANHO DA LOGO */
        height: 60px;
        z-index: 9999;
        /* background: transparent; */
        background: white;                  /* <--- FUNDO BRANCO SÓLIDO */
        /* ou background: rgba(255,255,255,0.8);  /* semi-transparente */ */
        padding: 15px;
        border-radius: 20px;
        box-shadow: 0 10px 35px rgba(0,0,0,0.6);
        text-align: center;
        pointer-events: none;
        overflow: hidden;
    ">
        <img src="data:image/png;base64,{encoded_string}" 
            alt="Logo" 
            style="
                width: 100%;
                height: 100%;
                object-fit: contain;
                border-radius: 15px;
            ">
    </div>
    '''

    m.get_root().html.add_child(folium.Element(logo_html))

    # Salva o mapa
    mapa_path = os.path.join("static", "mapa_votos.html")
    m.save(mapa_path)
    
    return render_template("mapa.html")
    # FIM

# PAGINA DE TESTE MATERIAIS

##@app.route("/teste_materiais", methods=["GET", "POST"])
##def teste_materiais():
    ##if request.method == "POST":
        ##print(f"Dados recebidos do formulário: {request.form.to_dict(flat=False)}")
        ##materiais = request.form.getlist("materiais_selecionados")  # Atualizado
        ##print(f"Materiais recebidos: {materiais}")
        ##flash(f"Materiais selecionados: {', '.join(materiais) if materiais else 'Nenhum'}", "success")
        ##return redirect(url_for("teste_materiais"))
    ##return render_template("teste_materiais.html")
    
    ##return render_template("questionario.html", 
                          ##alunos=alunos.to_dict('records'), 
                          ####nucleos=nucleos.to_dict('records'), 
                          ##bairros=bairros.to_dict('records'), 
                          ##professores=professores.to_dict('records'),
                          ##projetos=projetos.to_dict('records'),  # PLANILHA PROJETOS
                          ##turnos=turnos.to_dict('records'),
                          ##usuario=session.get("nome_usuario", ""))

# Rota para exportar questionários para Excel
@app.route("/exportar_questionarios")
def exportar_questionarios():
    questionarios = load_excel("Questionarios.xlsx")
    if questionarios.empty:
        questionarios = pd.DataFrame({"Mensagem": ["Nenhum Questionário Encontrado Para Exportação."]})
    
    os.makedirs("exportados", exist_ok=True)
    excel_path = "exportados/questionarios.xlsx"
    questionarios.to_excel(excel_path, index=False, engine='openpyxl')
    print(f"Planilha Salva Em: {excel_path}")
    return send_file(excel_path, as_attachment=True)

@app.route("/logout")
def logout():
    session.clear()  # Limpa a sessão
    flash("Você Saiu do App!", "success")
    return redirect(url_for("index"))  # linha alterada
    #return redirect(url_for("login"))  ## LINHA ORIGINAL

# Rota para exportar alunos para Excel
@app.route("/exportar_excel")
def exportar_excel():
    alunos = load_excel("Alunos.xlsx")
    amp = load_excel("Aluno_Modalidade_Polo.xlsx")
    modalidades = load_excel("Modalidades.xlsx")
    polos = load_excel("Polos.xlsx")
    
    df = alunos.merge(amp, left_on="Codigo", right_on="aluno_codigo", how="left")
    df = df.merge(modalidades, left_on="modalidade_codigo", right_on="Codigo", how="left", suffixes=("", "_mod"))
    df = df.merge(polos, left_on="polo_codigo", right_on="Codigo", how="left", suffixes=("", "_polo"))
    
    df = df[["Codigo", "Nome", "Data_nascimento", "Modalidade", "Nome_polo", "Endereco"]]
    if df.empty:
        df = pd.DataFrame({"Mensagem": ["Nenhum dado encontrado para exportação."]})
    
    os.makedirs("exportados", exist_ok=True)
    excel_path = "exportados/alunos_cadastrados.xlsx"
    df.to_excel(excel_path, index=False, engine='openpyxl')
    print(f"Planilha salva em: {excel_path}")
    return send_file(excel_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5001)