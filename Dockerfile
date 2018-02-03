FROM ruby:2.3.3
RUN apt-get update -qq && apt-get install -y build-essential libpq-dev nodejs
RUN mkdir /demo-app
WORKDIR /demo-app
COPY Gemfile /demo-app/Gemfile
COPY Gemfile.lock /demo-app/Gemfile.lock
RUN bundle install
COPY . /demo-app
