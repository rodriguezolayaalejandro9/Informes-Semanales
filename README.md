# Informe Semanal Automático

---

## 📌 Lo que hace
Este script genera y envía por correo electrónico informes semanales personalizados en formato PDF para cada estudiante del colegio.  
Los informes incluyen:
- Resumen de desempeños alcanzados y faltantes.
- Gráficos y tablas de seguimiento.
- Horario de clases semanal.

El sistema toma información de las bases de datos institucionales, organiza los datos por grado y asignatura, y envía los informes directamente al correo institucional de cada estudiante.

---

## 📚 Librerías usadas
- pandas  
- numpy  
- datetime  
- warnings  
- matplotlib  
- seaborn  
- smtplib  
- email.mime (multipart, text, application)  
- locale  

Además, requiere los módulos personalizados:
- `corregir_nombres`  
- `corregir_grados`  

---

## 📝 Notas
- Este código ahorró el trabajo **semanal** de **14 profesores**, cada uno de los cuales tardaba aproximadamente **1 hora** en realizar este proceso manualmente.
- Está diseñado para uso interno y adaptado a la estructura de datos del Colegio Gimnasio Kaiporé.
