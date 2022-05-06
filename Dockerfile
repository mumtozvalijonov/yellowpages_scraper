FROM python:3.8

RUN apt-get install -y wget

# Adding trusting keys to apt for repositories
RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add -

# Adding Google Chrome to the repositories
RUN sh -c 'echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list'
RUN sh -c 'echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google.list'

# Updating apt to see and install Google Chrome
RUN apt-get -y update

# Install Chrome
RUN apt-get install -y google-chrome-stable
RUN pip install --upgrade pip

#Creating User
COPY . /app 
WORKDIR /app
RUN useradd -ms /bin/bash newuser
RUN chown -R newuser:newuser /app
RUN chmod 755 /app
USER newuser



#Copy requirements and install
COPY ./requirements.txt /app/requirements.txt
RUN pip install -r requirements.txt