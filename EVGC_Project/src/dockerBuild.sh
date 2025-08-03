#!/usr/bin/bash

# Used for rebuilding the docker images after changes are made to the Dockerfile. 
# Run after I've made changes to the docker, or at least every day, especially if logging onto a different lab system for the first time.
# It takes about 3-4 minutes to run.

docker build -t docker-evgp -f Dockerfile .
