from flask import Flask, render_template, Response
import psycopg2 
import psycopg2.extras
import config 
import io
import xlwt 

app = Flask(__name__)

conn = psycopg2.connect(dbname=config.DB_NAME, user=config.DB_USER, password=config.DB_PASS, host=config.DB_HOST, port=config.DB_PORT)




@app.route("/")
def index():
  return render_template('index.html')

@app.route("/download/report/execl")
def download_report():
  cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
  cur.execute("select * from students order by id")
  result = cur.fetchall()
  # for row in result:
  #   print(row)

  # Output in bytes
  output = io.BytesIO()
  # Create WorkBook object
  workbook = xlwt.Workbook()
  # Add a Sheet 
  sh = workbook.add_sheet('Student Report')
  # Add headers 
  sh.write(0,0,'ID')
  sh.write(0,1,"First_Name")
  sh.write(0,2,"Last_Name")
  sh.write(0,3,"Email")

  idx = 0
  for row in result:
    sh.write(idx+1, 0, str(row['id']))
    sh.write(idx+1, 1, str(row['fname']))
    sh.write(idx+1, 2, str(row['lname']))
    sh.write(idx+1, 3, str(row['email']))
    idx += 1

  workbook.save(output)
  output.seek(0)

  return Response(output, mimetype="application/ms-excel",   headers={"Content-Disposition": "attachment;filename=student_report.xls"} )



  return render_template('index.html')


if __name__ == '__main__':
  app.run(debug=True)

