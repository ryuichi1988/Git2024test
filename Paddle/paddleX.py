from paddlex import create_pipeline

pipeline = create_pipeline(pipeline="table_recognition")
output = pipeline.predict(r'C:\Users\timaz\Documents\PythonFile\pd2\example6.png')
for res in output:
    res.print()
    res.save_to_img("./output/")
    res.save_to_json("./output/")