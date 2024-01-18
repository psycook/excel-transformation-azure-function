import azure.functions as func
from remap_blueprint import remap_blueprint

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)
app.register_functions(remap_blueprint)