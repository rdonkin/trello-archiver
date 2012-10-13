#!/usr/bin/env ruby
#
# Script to gather trello credentials
#
require 'yaml'

  LINK_TO_KEYS = "https://trello.com/1/appKey/generate"

  puts "Next we'll open a browser at a specific link to get the authentication keys"
  puts "Enter the name TrelloArchiver when it asks for the APP Name"
  `open #{LINK_TO_KEYS}`

  puts "Enter your key:"
  public_key = gets.chomp

  puts "Enter private key:"
  private_key = gets.chomp

  link_to_app_key = "https://trello.com/1/connect?key=#{public_key}&name=TrelloArchiver&response_type=token&scope=read,write,account&expiration=never"

  puts "Next we'll use that prior info to gather your application key"

  `open #{link_to_app_key}`

  access_token_key = gets.chomp

  puts "Dumping information into a personalized example config"
  puts "Read through content of config.personalized.yml"
  puts "When satisfied that it looks ok, then move it to config.yml"

  %w[public_key private_key access_token_key].each do |value|
    File.open('config.personalized.yml', 'a') { |f| f.write value.to_yaml }
  end
