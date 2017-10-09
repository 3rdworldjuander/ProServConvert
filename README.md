# Installing ccwparser  

## Requirements  
CentOS, python3, flask  

## Update system
$ sudo yum update  
$ sudo yum install yum-utils  
$ sudo yum groupinstall development  
$ sudo yum install epel-release  

## Install Python 3  
$ sudo yum install python34  

## Install pip  
$ curl -O https://bootstrap.pypa.io/get-pip.py  
$ sudo /usr/bin/python3.4 get-pip.py  

## Install virtualenv  
$ sudo pip3 install virtualenv  

## Clone from github  
$ git clone https://<copy link>  
$ cd ccwparser  
$ mkdir files  

## Activate virtualenv  
$ virtualenv -p /usr/bin/python3 venv  
$ source venv/bin/activate  
(venv)$ pip install flask  

## Install requirements.txt  
(venv)$ pip install -r requirements.txt  

## Test run  
(venv)$ python3 run.py  

## Configure for automatically running on server boot  
