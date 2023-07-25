from fastapi import FastAPI
from process import Vendas

app = FastAPI()


@app.get("/repor/{path_file}")
async def crete_report_from(path_file):
    print(path_file)
    path_file = str(path_file)
    vendas = Vendas("../temp/Relatorio-Vendas-js-82703.json")
    vendas.process_data()
    vendas.save_data()
    return {"message": path_file}
