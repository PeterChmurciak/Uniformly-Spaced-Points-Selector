# Uniformly-Spaced-Points-Selector
Python application with a simple GUI used for selecting a subset of uniformly spaced points.

# How to run

## (Optional) Create virtual envinronment 
###### Using virtualenv
```
virtualenv venv
.\venv\Scripts\activate
```
###### Using conda
``` 
conda create -n venv python
conda activate venv
```
## Install the required packages
```
pip install -r .\requirements.txt
```
## (Optional) Specify your API key for Google Maps
In file
```
.\Dependencies\Subroutine_Select_The_Points.py
```
on line 357, modify the variable
```
apikey = ''  # (your API key here)
```

## Run the main file
```
python Uniformly_Spaced_Points_Selector.py
```
## (Optional) Leave virtual envinronment 
###### Using virtualenv
```
deactivate
```
###### Using conda
``` 
conda deactivate
```
# Work in progress
Readme will be updated with explanation and pictures in the near future.
