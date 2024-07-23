from customtkinter import *

tela = CTk()
tela.geometry("500x400")

title = CTkLabel(tela, text="yago lindo do meu coração", font=("montserrat", 20)).pack(pady=20)
user = CTkEntry(tela, placeholder_text="Inserir Usuario...", width=300, font=("montserrat", 12)).pack(pady=20)
password = CTkEntry(tela, placeholder_text="Inserir senha...", width=300, show="*", font=("montserrat", 12)).pack()
checkbox = CTkCheckBox(tela, text="Lembrar de mim", font=("montserrat", 12)).place(x=100, y=180)
button = CTkButton(tela, text="Login", width=300, font=("montserrat", 20)).pack(pady=70)

tela.mainloop()
