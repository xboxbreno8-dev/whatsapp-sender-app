# -*- coding: utf-8 -*-
"""
WhatsApp Bulk Sender - Brenoxxx Edition
Aplicativo para envio de mensagens personalizadas via WhatsApp
"""

import os
import urllib.parse
import webbrowser
from threading import Thread
import time
import platform

# Imports do Kivy
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.popup import Popup
from kivy.uix.progressbar import ProgressBar
from kivy.uix.switch import Switch
from kivy.uix.scrollview import ScrollView
from kivy.clock import Clock
from kivy.core.window import Window

# Imports para manipulação de dados
try:
    import pandas as pd
except ImportError:
    pd = None

# Configuração da janela
Window.clearcolor = (0.1, 0.1, 0.15, 1)  # Cor de fundo escura


class FilePickerPopup(Popup):
    """Popup para seleção de arquivo Excel"""
    
    def __init__(self, callback, **kwargs):
        super().__init__(**kwargs)
        self.callback = callback
        self.title = "Selecionar Planilha Excel"
        self.size_hint = (0.9, 0.8)
        
        # Layout principal
        layout = BoxLayout(orientation='vertical', spacing=10, padding=10)
        
        # File chooser
        self.filechooser = FileChooserListView(
            filters=['*.xlsx', '*.xls'],
            path=os.path.expanduser('~')  # Começar na pasta home
        )
        layout.add_widget(self.filechooser)
        
        # Botões
        button_layout = BoxLayout(size_hint_y=None, height=50, spacing=10)
        
        select_btn = Button(
            text="Selecionar",
            background_color=(0.2, 0.7, 0.3, 1)
        )
        select_btn.bind(on_press=self.select_file)
        
        cancel_btn = Button(
            text="Cancelar",
            background_color=(0.7, 0.2, 0.2, 1)
        )
        cancel_btn.bind(on_press=self.dismiss)
        
        button_layout.add_widget(select_btn)
        button_layout.add_widget(cancel_btn)
        layout.add_widget(button_layout)
        
        self.content = layout
    
    def select_file(self, instance):
        """Selecionar arquivo e fechar popup"""
        if self.filechooser.selection:
            self.callback(self.filechooser.selection[0])
        self.dismiss()


class WhatsAppSender(BoxLayout):
    """Widget principal do aplicativo"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing = 15
        self.padding = [20, 20, 20, 20]
        
        # Variáveis de controle
        self.excel_file = None
        self.contacts_data = []
        self.current_index = 0
        self.auto_mode = False
        self.is_sending = False
        
        self.setup_ui()
    
    def setup_ui(self):
        """Configurar interface do usuário"""
        
        # Header com título personalizado
        header = BoxLayout(size_hint_y=None, height=80, spacing=10)
        
        # Logo/Ícone (usando texto estilizado)
        logo = Label(
            text="📱 WhatsApp\nBOT",
            font_size=18,
            halign='center',
            color=(0.2, 0.8, 0.4, 1),
            bold=True
        )
        
        # Título personalizado
        title = Label(
            text="Corretor Brenoxxx",
            font_size=16,
            halign='center',
            color=(1, 0.8, 0.2, 1),
            bold=True
        )
        
        header.add_widget(logo)
        header.add_widget(title)
        self.add_widget(header)
        
        # Área de mensagem
        msg_label = Label(
            text="📝 Digite sua mensagem personalizada:",
            size_hint_y=None,
            height=40,
            color=(1, 1, 1, 1),
            font_size=16,
            halign='left',
            bold=True
        )
        msg_label.text_size = (None, None)
        self.add_widget(msg_label)
        
        # Campo de texto para mensagem
        scroll_msg = ScrollView(size_hint_y=None, height=150)
        self.message_input = TextInput(
            multiline=True,
            hint_text="Digite aqui sua mensagem...\nUse {nome} para personalizar.\nExemplo: Olá {nome}! 😊\nComo vai você?",
            background_color=(0.2, 0.2, 0.25, 1),
            foreground_color=(1, 1, 1, 1),
            cursor_color=(0.2, 0.8, 0.4, 1),
            font_size=14,
            size_hint_y=None
        )
        self.message_input.bind(minimum_height=self.message_input.setter('height'))
        scroll_msg.add_widget(self.message_input)
        self.add_widget(scroll_msg)
        
        # Seletor de arquivo
        file_layout = BoxLayout(size_hint_y=None, height=60, spacing=10)
        
        file_label = Label(
            text="📊 Planilha:",
            size_hint_x=None,
            width=100,
            color=(1, 1, 1, 1),
            font_size=14
        )
        
        self.file_button = Button(
            text="Selecionar Planilha Excel",
            background_color=(0.3, 0.5, 0.8, 1),
            color=(1, 1, 1, 1)
        )
        self.file_button.bind(on_press=self.select_file)
        
        file_layout.add_widget(file_label)
        file_layout.add_widget(self.file_button)
        self.add_widget(file_layout)
        
        # Status da planilha
        self.file_status = Label(
            text="Nenhuma planilha selecionada",
            size_hint_y=None,
            height=30,
            color=(0.8, 0.8, 0.8, 1),
            font_size=12,
            halign='left'
        )
        self.file_status.text_size = (None, None)
        self.add_widget(self.file_status)
        
        # Controles de modo
        mode_layout = BoxLayout(size_hint_y=None, height=50, spacing=15)
        
        mode_label = Label(
            text="🎯 Modo:",
            size_hint_x=None,
            width=80,
            color=(1, 1, 1, 1),
            font_size=14
        )
        
        self.mode_switch = Switch(
            size_hint_x=None,
            width=100
        )
        self.mode_switch.bind(active=self.toggle_mode)
        
        self.mode_status = Label(
            text="Manual",
            color=(1, 0.8, 0.2, 1),
            font_size=14,
            bold=True
        )
        
        mode_layout.add_widget(mode_label)
        mode_layout.add_widget(self.mode_switch)
        mode_layout.add_widget(self.mode_status)
        self.add_widget(mode_layout)
        
        # Progresso
        progress_layout = BoxLayout(size_hint_y=None, height=40, spacing=10)
        
        progress_label = Label(
            text="Progresso:",
            size_hint_x=None,
            width=100,
            color=(1, 1, 1, 1),
            font_size=14
        )
        
        self.progress_bar = ProgressBar(
            max=100,
            value=0,
            size_hint_y=None,
            height=30
        )
        
        progress_layout.add_widget(progress_label)
        progress_layout.add_widget(self.progress_bar)
        self.add_widget(progress_layout)
        
        # Status atual
        self.status_label = Label(
            text="Pronto para começar",
            size_hint_y=None,
            height=30,
            color=(0.8, 0.8, 0.8, 1),
            font_size=12,
            halign='left'
        )
        self.status_label.text_size = (None, None)
        self.add_widget(self.status_label)
        
        # Botões de controle
        button_layout = BoxLayout(size_hint_y=None, height=60, spacing=10)
        
        self.send_button = Button(
            text="🚀 Iniciar Envios",
            background_color=(0.2, 0.7, 0.3, 1),
            color=(1, 1, 1, 1),
            font_size=16,
            bold=True
        )
        self.send_button.bind(on_press=self.start_sending)
        
        self.next_button = Button(
            text="⏭️ Próximo",
            background_color=(0.3, 0.5, 0.8, 1),
            color=(1, 1, 1, 1),
            disabled=True
        )
        self.next_button.bind(on_press=self.next_contact)
        
        self.stop_button = Button(
            text="⏹️ Parar",
            background_color=(0.7, 0.2, 0.2, 1),
            color=(1, 1, 1, 1),
            disabled=True
        )
        self.stop_button.bind(on_press=self.stop_sending)
        
        button_layout.add_widget(self.send_button)
        button_layout.add_widget(self.next_button)
        button_layout.add_widget(self.stop_button)
        self.add_widget(button_layout)
    
    def select_file(self, instance):
        """Abrir seletor de arquivo"""
        popup = FilePickerPopup(self.file_selected)
        popup.open()
    
    def file_selected(self, file_path):
        """Processar arquivo selecionado"""
        self.excel_file = file_path
        self.file_button.text = f"📊 {os.path.basename(file_path)}"
        
        # Tentar carregar e validar planilha
        try:
            self.load_contacts()
            self.file_status.text = f"✅ {len(self.contacts_data)} contatos carregados"
            self.file_status.color = (0.2, 0.8, 0.3, 1)
        except Exception as e:
            self.file_status.text = f"❌ Erro: {str(e)}"
            self.file_status.color = (0.8, 0.2, 0.2, 1)
    
    def load_contacts(self):
        """Carregar contatos da planilha"""
        if not pd:
            raise Exception("Pandas não está instalado")
        
        try:
            # Tentar diferentes engines para melhor compatibilidade
            if self.excel_file.endswith('.xlsx'):
                df = pd.read_excel(self.excel_file, engine='openpyxl')
            else:
                df = pd.read_excel(self.excel_file, engine='xlrd')
        except Exception as e:
            raise Exception(f"Erro ao ler arquivo: {str(e)}")
        
        df = pd.read_excel(self.excel_file)
        
        # Verificar colunas necessárias
        required_cols = ['Nome', 'Número de Telefone']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            # Tentar variações comuns dos nomes das colunas
            col_mapping = {
                'nome': 'Nome',
                'name': 'Nome',
                'numero': 'Número de Telefone',
                'telefone': 'Número de Telefone',
                'phone': 'Número de Telefone',
                'celular': 'Número de Telefone'
            }
            
            # Mapear colunas encontradas
            for col in df.columns:
                col_lower = col.lower().strip()
                if col_lower in col_mapping:
                    df.rename(columns={col: col_mapping[col_lower]}, inplace=True)
        
        # Verificar novamente
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise Exception(f"Colunas não encontradas: {', '.join(missing_cols)}")
        
        # Limpar e validar dados
        df = df.dropna(subset=['Nome', 'Número de Telefone'])
        df['Nome'] = df['Nome'].astype(str).str.strip()
        df['Número de Telefone'] = df['Número de Telefone'].astype(str).str.strip()
        
        # Limpar números de telefone (remover caracteres especiais)
        df['Número de Telefone'] = df['Número de Telefone'].str.replace(r'[^\d+]', '', regex=True)
        
        self.contacts_data = df.to_dict('records')
        self.current_index = 0
        
        if not self.contacts_data:
            raise Exception("Nenhum contato válido encontrado na planilha")
    
    def toggle_mode(self, instance, value):
        """Alternar entre modo manual e automático"""
        self.auto_mode = value
        if value:
            self.mode_status.text = "Automático"
            self.mode_status.color = (0.2, 0.8, 0.3, 1)
        else:
            self.mode_status.text = "Manual"
            self.mode_status.color = (1, 0.8, 0.2, 1)
    
    def start_sending(self, instance):
        """Iniciar processo de envio"""
        if not self.validate_inputs():
            return
        
        self.is_sending = True
        self.current_index = 0
        
        # Atualizar interface
        self.send_button.disabled = True
        self.next_button.disabled = False
        self.stop_button.disabled = False
        
        # Processar primeiro contato
        self.process_current_contact()
    
    def validate_inputs(self):
        """Validar entradas do usuário"""
        if not self.message_input.text.strip():
            self.show_error("Por favor, digite uma mensagem")
            return False
        
        if not self.contacts_data:
            self.show_error("Por favor, selecione uma planilha válida")
            return False
        
        return True
    
    def process_current_contact(self):
        """Processar contato atual"""
        if self.current_index >= len(self.contacts_data):
            self.finish_sending()
            return
        
        contact = self.contacts_data[self.current_index]
        
        # Personalizar mensagem
        message = self.message_input.text
        message = message.replace('{nome}', contact['Nome'])
        
        # Codificar para URL
        encoded_message = urllib.parse.quote(message, safe='')
        
        # Limpar número de telefone
        phone = str(contact['Número de Telefone']).strip()
        if not phone.startswith('+'):
            phone = '+55' + phone  # Assumir Brasil se não tiver código do país
        
        # Gerar URL do WhatsApp
        whatsapp_url = f"https://wa.me/{phone.replace('+', '')}?text={encoded_message}"
        
        # Atualizar status
        progress = ((self.current_index + 1) / len(self.contacts_data)) * 100
        self.progress_bar.value = progress
        self.status_label.text = f"Enviando para: {contact['Nome']} ({phone})"
        
        try:
            # Abrir WhatsApp
            webbrowser.open(whatsapp_url)
            
            # Se modo automático, aguardar e continuar
            if self.auto_mode and self.is_sending:
                Clock.schedule_once(self.auto_next, 5)  # 5 segundos de delay
            
        except Exception as e:
            self.show_error(f"Erro ao abrir WhatsApp para {contact['Nome']}: {str(e)}")
    
    def auto_next(self, dt):
        """Avançar automaticamente para próximo contato"""
        if self.is_sending:
            self.next_contact(None)
    
    def next_contact(self, instance):
        """Avançar para próximo contato"""
        if not self.is_sending:
            return
        
        self.current_index += 1
        
        if self.current_index < len(self.contacts_data):
            self.process_current_contact()
        else:
            self.finish_sending()
    
    def stop_sending(self, instance):
        """Parar processo de envio"""
        self.is_sending = False
        
        # Restaurar interface
        self.send_button.disabled = False
        self.next_button.disabled = True
        self.stop_button.disabled = True
        
        self.status_label.text = f"Parado. {self.current_index} de {len(self.contacts_data)} mensagens enviadas"
    
    def finish_sending(self):
        """Finalizar processo de envio"""
        self.is_sending = False
        
        # Restaurar interface
        self.send_button.disabled = False
        self.next_button.disabled = True
        self.stop_button.disabled = True
        
        self.progress_bar.value = 100
        self.status_label.text = f"✅ Concluído! {len(self.contacts_data)} mensagens enviadas"
        
        # Mostrar popup de conclusão
        self.show_success(f"Todas as {len(self.contacts_data)} mensagens foram processadas!")
    
    def show_error(self, message):
        """Mostrar popup de erro"""
        popup = Popup(
            title="❌ Erro",
            content=Label(text=message, text_size=(300, None), halign='center'),
            size_hint=(0.8, 0.4)
        )
        popup.open()
    
    def show_success(self, message):
        """Mostrar popup de sucesso"""
        popup = Popup(
            title="✅ Sucesso",
            content=Label(text=message, text_size=(300, None), halign='center'),
            size_hint=(0.8, 0.4)
        )
        popup.open()


class WhatsAppBulkSenderApp(App):
    """Aplicativo principal"""
    
    def build(self):
        self.title = "WhatsApp Bulk Sender - Brenoxxx Edition"
        return WhatsAppSender()


if __name__ == '__main__':
    WhatsAppBulkSenderApp().run()