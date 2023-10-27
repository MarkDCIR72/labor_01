class WidgetGridConfigurer:
    def __init__(self, frame_):
        self.frame = frame_

    def configure_widgets(self):
        for widget in self.frame.winfo_children():
            widget.grid_configure(padx=10, pady=5)
