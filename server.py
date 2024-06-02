from flask import Flask, jsonify
import win32com.client
import pythoncom

app = Flask(__name__)

def move_slide(index):
    try:
        pythoncom.CoInitialize()
        
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        
        if ppt_app.Presentations.Count > 0:
            presentation = ppt_app.ActivePresentation
            
            next_slide_index = presentation.SlideShowWindow.View.Slide.SlideIndex + index
            
            if next_slide_index <= presentation.Slides.Count:
                presentation.SlideShowWindow.View.GotoSlide(next_slide_index)
                result = {"status": "success", "message": f"Moved to slide {next_slide_index}"}
            else:
                result = {"status": "info", "message": "Already on the last slide."}
        else:
            result = {"status": "error", "message": "No presentations are currently open."}
        
        pythoncom.CoUninitialize()
        
        return result
    except Exception as e:
        pythoncom.CoUninitialize()
        return {"status": "error", "message": str(e)}

@app.route('/ping', methods=['GET'])
def hello():
    return 'pong'

@app.route('/move_forward', methods=['POST'])
def move_forward():
    print("move forward")
    result = move_slide(1)
    return jsonify(result)

@app.route('/move_backword', methods=['POST'])
def move_backwords():
    print("move backword")
    result = move_slide(-1)
    return jsonify(result)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
