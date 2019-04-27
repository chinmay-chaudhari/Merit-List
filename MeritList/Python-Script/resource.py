from flask_restful import Resource, Api
from flask import Flask, request
import subprocess as p

app = Flask(__name__)
api = Api(app)

graph_array=[]

class fileRun(Resource):
    def put(self):
        global graph_array
        data=request.get_json()
        fname=str(data['fname'])
        if fname == "bsc.xlsx":
            graph_array = p.check_output("python3 Bsc.py "+fname, shell=True).decode("utf-8")
        if fname == "msc.xlsx":
            graph_array = p.check_output("python3 Msc.py "+fname, shell=True).decode("utf-8")
        return 1
class getGraphs(Resource):
    def get(self):
        global graph_array
        arr = graph_array.rstrip().split(',')
        graph_array=[]
        return {'names':arr}

api.add_resource(getGraphs,"/getGraphs")
api.add_resource(fileRun,"/fileRun")

if __name__ == '__main__':
     app.run(host='127.0.0.1')
