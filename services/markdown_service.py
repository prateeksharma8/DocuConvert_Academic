from markdown import markdown


def parse_markdown_to_html(markdown_text: str, title: str, template: str) -> str:
    rendered = markdown(markdown_text)
    template_name = template if template in {"basic", "report"} else "basic"

    if template_name == "report":
        return f"""
        <html>
        <head>
            <meta charset="utf-8" />
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; color: #1f2937; line-height: 1.7; }}
                h1, h2, h3 {{ color: #111827; }}
                img {{ max-width: 100%; height: auto; margin: 18px 0; border-radius: 8px; }}
                .container {{ max-width: 900px; margin: 0 auto; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>{title}</h1>
                {rendered}
            </div>
        </body>
        </html>
        """

    return f"""
    <html>
    <head>
        <meta charset="utf-8" />
        <style>
            body {{ font-family: sans-serif; margin: 40px; color: #111827; line-height: 1.6; }}
            h1, h2, h3 {{ color: #111827; }}
            img {{ max-width: 100%; height: auto; margin: 16px 0; }}
            .container {{ max-width: 800px; margin: 0 auto; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>{title}</h1>
            {rendered}
        </div>
    </body>
    </html>
    """
