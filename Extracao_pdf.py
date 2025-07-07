from botcity.core import DesktopBot
import os
import tkinter as tk
from tkinter import simpledialog, messagebox
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials

# Escopo necessário para acesso ao Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ⚠️ Substitua pelo ID da sua planilha no ambiente local (nunca publique com o valor real)
ID_PLANILHA = "SUA_PLANILHA_AQUI"  # <- configurar localmente
RANGE_COLUNA = "ENVIO DE PDFS!A2:A"

class Bot(DesktopBot):

    # Autentica com a API do Google Sheets (client_secret.json deve estar localmente)
    def autenticar_google_sheets(self):
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                # ⚠️ Este arquivo precisa estar presente apenas localmente
                flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
                creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return build("sheets", "v4", credentials=creds)

    # Lê os nomes dos clientes da planilha
    def obter_dados_planilha(self):
        service = self.autenticar_google_sheets()
        result = service.spreadsheets().values().get(
            spreadsheetId=ID_PLANILHA,
            range=RANGE_COLUNA
        ).execute()
        dados = result.get("values", [])
        return [linha[0] for linha in dados if linha and linha[0].strip()]

    # Marca a linha como "OK" na coluna B
    def marcar_processado(self, linha):
        service = self.autenticar_google_sheets()
        body = {"values": [["OK"]]}
        service.spreadsheets().values().update(
            spreadsheetId=ID_PLANILHA,
            range=f"ENVIO DE PDFS!B{linha}",
            valueInputOption="RAW",
            body=body,
        ).execute()

    def action(self, execution=None):
        self.realizar_login()

        while True:
            clientes = self.obter_dados_planilha()
            if not clientes:
                messagebox.showinfo("Pronto", "Todos os clientes processados.")
                break

            nome_tecnico = (
                "Equipe Sebratel"
                if messagebox.askquestion("Seleção", "Equipe Sebratel?") == "yes"
                else "Terceiros"
            )

            for idx, nome in enumerate(clientes, start=2):
                self.processar_cliente(nome, nome_tecnico)
                self.marcar_processado(idx)

    def realizar_login(self):
        self.browse("https://erp.sebratel.net.br/users/login")
        self.wait(5000)
        self.tab()
        self.wait(3000)

        # ⚠️ Substituir pelas credenciais corretas no ambiente local
        self.paste('SEU_USUARIO_AQUI')
        self.tab()
        self.wait(1000)
        self.paste('SUA_SENHA_AQUI')
        self.wait(1000)
        self.enter()
        self.wait(3000)
        self.control_t()
        self.paste('https://erp.sebratel.net.br/service_dashboard#list/team')
        self.enter()

    def processar_cliente(self, nome, nome_tecnico):
        nome_formatado = nome.replace('/', '-')
        self.buscar_cliente(nome)
        self.gerar_pdf(nome_formatado, nome_tecnico)
        self.salvar_pdf(nome_formatado, nome_tecnico)

    def buscar_cliente(self, nome):
        if self.find("PESQUISAR", matching=0.97, waiting_time=10000):
            self.click()
            self.paste(nome)
            self.wait(6000)
        else:
            self.not_found("PESQUISAR")

    def gerar_pdf(self, nome, nome_tecnico):
        if self.find("CLIENTE", matching=0.97, waiting_time=10000):
            self.click_relative(60, 227)
        if self.find("PDF", matching=0.97, waiting_time=10000):
            self.click()
        if self.find("TEMPLATE", matching=0.97, waiting_time=10000):
            self.click_relative(100, 64)

        self.wait(10000)
        self.enter()
        self.wait(5000)
        self.paste(nome)
        self.enter()
        self.wait(5000)
        self.control_w()

    def salvar_pdf(self, nome, nome_tecnico):
        caminho_origem = r'C:\Users\SEU_USUARIO\Desktop'
        caminho_destino = rf'C:\Users\SEU_USUARIO\Documents\{nome_tecnico}'

        if not os.path.exists(caminho_destino):
            os.makedirs(caminho_destino)

        for arquivo in os.listdir(caminho_origem):
            if arquivo.endswith('.pdf'):
                origem = os.path.join(caminho_origem, arquivo)
                destino = os.path.join(caminho_destino, f"{nome}.pdf")
                try:
                    os.rename(origem, destino)
                    print(f"PDF salvo: {destino}")
                except Exception as e:
                    print(f"Erro ao mover PDF: {e}")

        if self.find("Minhas3", matching=0.97, waiting_time=10000):
            self.triple_click_relative(73, 49)
            self.delete()
        else:
            self.not_found("Minhas3")
            print('CAMPO não encontrado!')

    def not_found(self, elemento):
        print(f"Elemento não encontrado: {elemento}")


if __name__ == "__main__":
    Bot.main()
