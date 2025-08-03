# Docker Instructions

## Quick Start

Run "./dockerBuild.sh" to build the image, especially if changes have been made to the container recently, or if you've logged onto a lab system that you haven't built an image on before.

Run "./dockerRun.sh" to launch docker container. You will be placed in the "/evgp" folder and have full access to our git repo. You will be in the default main branch.


## Using git with Docker

git in the "/evgp" folder works like git outside of the Docker. You can switch to the desired remote branch with "git switch <branch>", 
and run "git add <file>"
        "git commit"
        "git pull"
        "git push" 

to commit your changes to that branch.


## Installed Software

The Docker container has:

- python3
- python3-pip, 
- python3-virtualenv
- python-pptx 
- python-pptx-interface 
- flask
- pillow
- vim
- libreoffice
- imagemagick
- curl
