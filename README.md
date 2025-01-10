# Eyemark Backend

The Eyemark Backend Codebase is built with Django and Django-rest framework. The database used in development and production is PostgreSQL. 
Celery is set up with Redis for running and processing background tasks. There's a docker set up for different environment types (local, staging, production) all you have to do is run the correct startup command for what environment you are running on.

Fetching the application:
```bash 
$ git clone git@github.com:zstechnology/eyemark_backend.git
```

Go to the project root directory
```bash
$ cd eyemark_backend
```

# Environment Variables
The .envs folder should contain the necessary environment variables for building and running the application, take a sample from .env.example and populate them.

Once you have populated the neccesary environment variables, the directory structure for the .env folder should follow as below

`
|-- .envs
    |-- .local
      |-- .django
      |-- .postgres

`

# Building and running the project

To run the project locally, you have to have docker and docker compose installed.

At the root of the directory run the following commands respectively to build and run the project:
```bash
$ docker-compose -f local.yml build
```

```bash
$ docker-compose -f local.yml up
```
