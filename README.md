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
To run the project you will need to add the following environment variables to your django env file in the .envs folder
`USE_DOCKER`
`IPYTHONDIR`

`REDIS_URL`

`CELERY_FLOWER_USER`
`CELERY_FLOWER_PASSWORD`

`OTP_SECRET_KEY`

`STREAM_API_KEY`
`STREAM_API_SECRET`
`DJANGO_SECRET_KEY`
`DJANGO_ADMIN_URL`

`EMAIL_HOST`
`EMAIL_HOST_USER`
`SENDGRID_API_KEY`
`BREVO_API_KEY`

`CLOUDINARY_CLOUD_NAME`
`CLOUDINARY_API_KEY`
`CLOUDINARY_API_SECRET`
`TEST_UPLOAD`

`STATE_URL`
`LGA_URL`
`WARD_URL`
`AREA_URL`
`PRIMARY_SCHOOL_URL`
`SECONDARY_SCHOOL_URL`
`TERTIARY_SCHOOL_URL`
`HEALTH_CARE_URL`
`CHURCH_URL`
`MOSQUE_URL`
`ROAD_URL`

`WATER_POINT_URL`
`FRONTEND_URL`
`GOOGLE_APPLICATION_CREDENTIALS`
`ACCESS_TOKEN_LIFETIME`
`MAPS_API_KEY`
`BOUNTY_API_SECRET`
`BOUNTY_DATABASE_ID`

`AWS_ACCESS_KEY`
`AWS_ACCESS_KEY_ID`
`AWS_BUCKET_NAME`
`AWS_REGION`

To run the project locally, you have to have docker and docker compose installed.

At the root of the directory run the following commands respectively to build and run the project:
```bash
$ docker-compose -f local.yml build
```

```bash
$ docker-compose -f local.yml up
```
