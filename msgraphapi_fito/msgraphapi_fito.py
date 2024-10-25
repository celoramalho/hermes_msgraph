#!/usr/bin/env python
# coding: utf-8
import json
import requests
import yaml
import pandas as pd

class MSGraphAPI:
    """
    Classe para interagir com a API do Microsoft Graph para envio e leitura de e-mails.

    A classe contém métodos para obter o token de acesso, enviar e-mails, ler mensagens de e-mail, organizar os dados em um DataFrame e salvar em arquivo JSON.
    """

    def __init__(self, client_id, client_secret, tenant_id):
        """
        Inicializa a classe com as credenciais passadas como parametro
        client_id, client_secret, tenant_id
        ----------
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id

    def __get_access_token(self):
        """
        Método privado para obter o token de acesso à API do Microsoft Graph.

        Retorna:
        -------
        access_token : str
            Token de acesso para realizar chamadas à API do Microsoft Graph.
        """
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default"
        }
        try:
            response = requests.post(url, data=payload)
        except Exception as e:
            raise RuntimeError("Não foi possível retornar o request do post para pegar o access token") from e
        
        response_data = response.json()
        access_token = response_data['access_token']
        return access_token

    def send_email(self, sender_mail, subject, body, to_address, cc_address=None):
        """
        Envia um e-mail usando a API do Microsoft Graph.

        Parameters:
        ----------
        sender_mail : str
            E-mail do remetente.
        subject : str
            Assunto do e-mail.
        body : str
            Corpo do e-mail.
        to_address : str
            Endereço de e-mail do destinatário.
        cc_address : str, optional
            Endereço de e-mail para cópia, por padrão None.
        """
        access_token = self.__get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        url = f"https://graph.microsoft.com/v1.0/users/{sender_mail}/sendMail"

        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_address
                        }
                    }
                ],
                "ccRecipients": [
                    {
                        "emailAddress": {
                            "address": cc_address
                        }
                    }
                ] if cc_address else []
            },
            "saveToSentItems": "true"
        }

        response = requests.post(url, headers=headers, data=json.dumps(payload))

        if response.status_code == 200:
            print("E-mail enviado com sucesso!")
        else:
            print(f"Erro ao enviar e-mail: {response.status_code}")

    def __read_email(self, messages_json_path, email_address, n_of_massages):
        """
        Método privado para ler e-mails de um endereço específico.

        Parameters:
        ----------
        messages_json_path : str
            Caminho do arquivo JSON onde as mensagens serão salvas.
        email_address : str
            Endereço de e-mail do qual ler as mensagens.
        n_of_massages : int
            Número de mensagens a serem recuperadas.
        """
        access_token = self.__get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        url = f"https://graph.microsoft.com/v1.0/users/{email_address}/messages?$top={n_of_massages}"
        
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            messages = response.json().get('value', [])
            try:
                with open(messages_json_path, 'w+', encoding='utf-8') as file:
                    json.dump(messages, file, ensure_ascii=False, indent=4)
            except Exception as e:
                raise RuntimeError("Não foi possível salvar as mensagens no arquivo messages.json") from e
        else:
            print(f"Erro ao ler e-mails: {response.status_code}")

    def __json_to_dataframe(self, json_file_path):
        """
        Método privado para converter um arquivo JSON de mensagens de e-mail em um DataFrame.

        Parameters:
        ----------
        json_file_path : str
            Caminho do arquivo JSON com as mensagens de e-mail.

        Returns:
        -------
        df_messages : pandas.DataFrame
            DataFrame contendo as mensagens de e-mail.
        """
        with open(json_file_path, 'r') as file:
            json_data = json.load(file)
        df_messages = pd.json_normalize(json_data)
        return df_messages

    def __organize_df_messages(self, df_messages):
        """
        Método privado para organizar o DataFrame de mensagens de e-mail com colunas relevantes.

        Parameters:
        ----------
        df_messages : pandas.DataFrame
            DataFrame com as mensagens de e-mail.

        Returns:
        -------
        df_emails : pandas.DataFrame
            DataFrame organizado com colunas específicas.
        """
        columns = ['subject', 'isRead', 'sentDateTime', 'receivedDateTime', 'sender.emailAddress.name', 
                   'sender.emailAddress.address', 'from.emailAddress.name', 'from.emailAddress.address', 
                   'bodyPreview', 'body.contentType', 'body.content']
        df_emails = df_messages[columns]
        return df_emails

    def get_df_emails(self, email_address, n_of_massages=10, messages_json_path='messages.json'):
        """
        Método público para obter as mensagens de e-mail e organizá-las em um DataFrame apenas com colunas normalmente necessárias, sem todos os atributos do e-mail.

        Parameters:
        ----------
        email_address : str
            Endereço de e-mail do qual obter as mensagens.
        messages_json_path : str, optional
            Caminho do arquivo JSON onde os e-mails serão salvas, por padrão 'messages.json'.
        n_of_massages : int, optional
            Número de mensagens a coletar, por padrão 10.

        Returns:
        -------
        df_messages : pandas.DataFrame
            DataFrame organizado com as mensagens de e-mail.
        """
        self.__read_email(messages_json_path, email_address, n_of_massages)
        df_full_messages = self.__json_to_dataframe(messages_json_path)
        df_messages = self.__organize_df_messages(df_full_messages)
        return df_messages
    # Exemplo de uso:
    # api = MSGraphAPI()
    # df_messages = api.get_df_emails("anakin.skywalker@fitenergia.com.br")

    def get_raw_df_emails(self, email_address, n_of_massages=10, messages_json_path='messages.json'):
        """
        Método público para obter as mensagens de e-mail e organizá-las em um DataFrame com todos os atributos recebidos.

        Parameters:
        ----------
        email_address : str
            Endereço de e-mail do qual obter as mensagens.
        messages_json_path : str, optional
            Caminho do arquivo JSON onde as mensagens serão salvas, por padrão 'messages.json'.
        n_of_massages : int, optional
            Número de e-mails que deseja ver, por padrão 10.

        Returns:
        -------
        df_messages : pandas.DataFrame
            DataFrame completo com as mensagens de e-mail e todos os atributos.
        """
        self.__read_email(messages_json_path, email_address, n_of_massages)
        df_full_messages = self.__json_to_dataframe(messages_json_path)
        return df_full_messages