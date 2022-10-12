FROM public.ecr.aws/lambda/python:3.9
RUN yum update -y && yum install zip -y
RUN mkdir -p /dist
COPY ./lambda_function.py /lambda_function.py
COPY ./requirements.txt /requirements.txt
RUN pip install --target /package -r /requirements.txt
WORKDIR /package
RUN zip -r9 /function.zip .
WORKDIR /
ENTRYPOINT []