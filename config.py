apikey = ""
default_model = "gpt-4o-mini-2024-07-18"
class Config:
    def __init__(self, open_api_key=None, model=None):
        if open_api_key:
            self.open_api_key = open_api_key
        else:
            self.open_api_key = apikey

        if model:
            self.model = model
        else:
            self.model = default_model