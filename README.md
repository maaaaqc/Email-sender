# Python server for generating cd-test summary in excel

## Summary
The server converts json results of the cd-test into a merged excel file and sends the excel file to the email that the user posts.
The cd-test folder is not needed locally as the server will create a temporary directory to clone cd-test from gitlab and cleanup after.

## Routes
The server provides 2 routes.

* /custom to customize the branches and folders that the cd test runs in
* /default to run cd tests using the branches and folders in a default config file

### Example
Start the container:
```console
docker build -t summary .
docker run --name summary -p 7710:7710 -v /home/qingchuan/cd-test-summary/config.json:/cd-summary/config.json:ro summary
```

Using curl command:
```console
curl -i -X POST 'http://localhost:7710/default' -H 'content-type: application/json' -d '{"filename": "server1", "email": "qingchuan.ma@nuance.com"}'
```
