# Excel add-in with xlwings and AWS Lambda (Python)

This project shows you how to implement xlwings Server with AWS Lambda functions.  
Make sure to check out the [xlwings Server docs](https://docs.xlwings.org/en/stable/remote_interpreter.html#remote-interpreter) for more in-depth information.

## AWS Lambda

1. Create the Lambda function

    * Open the [Lambda console](https://console.aws.amazon.com/lambda)
    * Click on `Create function` on the top right
    * Function name: Enter a function name, e.g. `xlwings-test`
    * Runtime: Choose `Python 3.9` in the dropdown
    * Expand `Advanced settings`, then activate the checkbox `Enable function URL`
    * Click on `Create function` at the bottom right

2. Configure the Lambda function

    Clicking the `Create function` from the previous step will open the detailed page with your Lambda function. In the middle of the page, click on the `Configuration` tab as you need to configure 2 environment variables. In the left-hand menu, click on `Environment variables` > `Edit` > `Add environment variable`. Add the following two environment variables:

    ```
    Key: XLWINGS_LICENSE_KEY
    Value: yourkey
    ```

    Note that you can get an xlwings PRO trial key at [https://www.xlwings.org/trial](https://www.xlwings.org/trial)

    and 

    ```
    Key: XLWINGS_API_KEY
    Value: DEVELOPMENT
    ```

    For both environment variables, hit `Save`.

# Create and upload function.zip

If you use 3rd party dependencies like pandas or xlwings, you will need to provide AWS Lambda with either a ZIP file (`function.zip`) or a Docker image. Since the Docker option requires you to use Amazon ECR container registry, we're going to upload a ZIP file instead. And as AWS Lambda will require specific versions of these dependencies that match the Lambda runtime environment, it's easiest to build the ZIP file in a Docker container with the same environment.

Therefore, to create our `function.zip` with pandas and xlwings as dependencies, run the following command (you'll need a recent version of Docker):

```
docker compose run --rm app
```

This will create the file `functions.zip` under the `dist` folder. Note that if you change your dependencies under `requirements.txt`, you'll need to rebuild your image as follows:

```
docker compose build
```

Once rebuilt, you can run the previous command again to produce the `function.zip` file. To recreate the ZIP file with only code changes, you don't need the build step and can run `docker compose run --rm app` again.

To upload the  ZIP file, go the the `Code` tab on the AWS Lambda page of your function and click on the `Upload from` button on the right hand side. Select `.zip File` in the dropdown, click on `Upload` and browse to your `function.zip`. Confirm with `Open`, and finally hit `Save`. Once uploaded, the function is available at the `Function URL` shown in the box at the top of the page.

Instead of uploading your ZIP file manually, you could also use the [AWS CLI](https://docs.aws.amazon.com/cli/latest/reference/lambda/update-function-code.html). 

## Excel add-in

Open `demo.xlam` by double-clicking it (**make sure to enable macros**). You should see a new Ribbon tab called `Demo`. Open the VBA editor via `Alt+F11` (Windows) or `Option-F11` (macOS), then, in the module `RibbonDemo`, change `"URL"` at the top of the module to the `Function URL` as shown in AWS Lambda.

Now create a new blank workbook and Fill in some data in cell `A1:B2`:

```
  | A | B |
-----------
1 | a | b |
-----------
2 | 1 | 2 |
```

Then hit the `Run` button on the add-in. It should write back the transposed version into cell `G1`.
