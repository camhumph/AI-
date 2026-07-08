@echo off
cd /d C:\CMS_AI
call venv\Scripts\activate
python C:\CMS_AI\yolo_box_editor.py --image "C:\CMS_AI\datasets\major8_full_views\images\train\Dynacast-2223488-C18595.sldasm_FULLTRAIN_RIGHT.jpg" --label "C:\CMS_AI\datasets\major8_full_views\labels\train\Dynacast-2223488-C18595.sldasm_FULLTRAIN_RIGHT.txt" --data "C:\CMS_AI\datasets\major8_full_views\data.yaml"
