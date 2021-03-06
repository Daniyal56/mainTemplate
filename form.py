from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SubmitField, BooleanField
from wtforms.validators import DataRequired, Length, Email, EqualTo


class Form(FlaskForm):
    invoice = StringField('invoice',
                          validators=[DataRequired()])
    submit = SubmitField('Submit')
