# CI/CD Documentation

Steps of our CI/CD pipeline

1. You push changes to a branch.
2. The commit triggers a git hook.
3. The git hook starts a Gitlab runner, which is a machine that runs automated Gitlab tests. Amdahl is our runner.
4. The Gitlab runner looks for a ".gitlab-ci.yml" file in the branch, and runs the file if found.
5. The runner starts a Docker container on Amdahl with Docker-in-Docker enabled, and copies our git repo to the runner.
6. The runner builds and starts our Docker container, using Dockerfile_CICD inside of its own Docker container.
7. When our Docker container runs, it automatically starts the Flask web server from app.py, as our Dockerfile tells it to run app.py as soon as the container starts.
8. The runner pings the Flask web server in our Docker container using curl, which is installed on its own Docker container.
