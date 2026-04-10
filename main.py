from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/generate")
def generate_proposal(data: dict):
    return {
        "status": "success",
        "message": "Payload recebido com sucesso",
        "data_received": data
    }
