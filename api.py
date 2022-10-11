from fastapi import FastAPI, Body, HTTPException

app = FastAPI()


@app.get("/")
async def root():
    return {'status': 'OK'}


@app.post('/createfile')
async def createfile(payload: dict = Body(...)):
    if 'slides' in payload:
        return {'msg': 'file', 'payload': payload}
    else:
        raise HTTPException(500, title='no slides array ..')
