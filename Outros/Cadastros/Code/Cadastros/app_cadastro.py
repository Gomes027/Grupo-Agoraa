from Cadastros.interface_usuario import InterfaceUsuario

class AppCadastros:
    def __init__(self, on_close=None):
        self.on_close = on_close
        self.ui = InterfaceUsuario(self)

    def run(self):
        self.ui.mainloop()

    def close(self):
        if self.on_close:
            self.on_close()
        self.ui.destroy()