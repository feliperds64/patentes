from estevam_flask import db
from datetime import datetime

class Dados(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    prop_ter = db.Column(db.String, nullable=False)
    cod_processo = db.Column(db.Integer, nullable=False)
    nom_marca = db.Column(db.String, nullable=False)
    nom_titular = db.Column(db.String, nullable=False)
    desc_desp = db.Column(db.String, nullable=False)
    classe = db.Column(db.Integer, nullable=False)
    especificacao = db.Column(db.String, nullable=False)
    imagem = db.Column(db.String, default = 'default.jpg', nullable=False)
    data_referencia = db.Column(db.DateTime, nullable=False)