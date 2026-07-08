@echo off
cd /d C:\CMS_AI
call venv\Scripts\activate

echo ==================================================
echo Started major8 full views training: %DATE% %TIME%
echo Dataset: C:\CMS_AI\datasets\major8_full_views\data.yaml
echo Starting model: C:\CMS_AI\models\cms_mold_right_view8_overnight.pt
echo ==================================================

yolo detect train model=C:\CMS_AI\models\cms_mold_right_view8_overnight.pt data=C:\CMS_AI\datasets\major8_full_views\data.yaml epochs=300 imgsz=1280 batch=1 device=cpu mosaic=0 fliplr=0 translate=0 scale=0 hsv_h=0 hsv_s=0 hsv_v=0 erasing=0 cls=3.0 patience=300 name=major8_full_views_1280_morebases
if errorlevel 1 (
  echo Training failed with error %ERRORLEVEL%
  exit /b %ERRORLEVEL%
)

echo ==================================================
echo Training finished: %DATE% %TIME%
echo Copying best model...
echo ==================================================

copy /Y C:\CMS_AI\runs\detect\major8_full_views_1280_morebases\weights\best.pt C:\CMS_AI\models\cms_mold_major8_full_views.pt
if errorlevel 1 (
  echo Copy failed with error %ERRORLEVEL%
  exit /b %ERRORLEVEL%
)

echo ==================================================
echo Running quick prediction on new left image...
echo ==================================================

yolo predict model=C:\CMS_AI\models\cms_mold_major8_full_views.pt source="C:\CMS_AI\datasets\major8_full_views\images\train\26036-37_FRAME_ONLY_20260508_FULLTRAIN_LEFT.jpg" imgsz=1280 conf=0.10 save=True name=major8_full_views_left_test

echo ==================================================
echo All done: %DATE% %TIME%
echo Final model: C:\CMS_AI\models\cms_mold_major8_full_views.pt
echo ==================================================
