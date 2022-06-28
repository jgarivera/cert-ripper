class TemplateReader:
    def __init__(self, template_path: str) -> None:
        self.template_path = template_path

    def read(self) -> str:
        print(f"Using template for email: \x1b[94m{self.template_path}\x1b[0m")

        with open(self.template_path) as file:
            template_html = file.read()

        return template_html
