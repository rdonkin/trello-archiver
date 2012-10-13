# encoding: utf-8
#!/usr/bin/env ruby

require 'trello'
require 'rubygems'
require_relative '../lib/trello-archiver'
require 'yaml'

include Trello
include Trello::Authorization

Trello::Authorization.const_set :AuthPolicy, OAuthPolicy

CONFIG = YAML::load(File.open("config.yml")) unless defined? CONFIG

credential = OAuthCredential.new CONFIG['public_key'], CONFIG['private_key']
OAuthPolicy.consumer_credential = credential

OAuthPolicy.token = OAuthCredential.new CONFIG['access_token_key'], nil

ignore = CONFIG['ignore']

me = Member.find("me")
boardarray = Array.new
me.boards.each do |board|
  if ignore.include? board.id
	  puts "Skipping #{board.name}"
  else
    filename = board.name.parameterize
    puts "Preparing to backup #{board.name}"
    TrelloArchiver::Archiver.new(:board => board, :filename => filename, :format => 'csv').create_backup
	end
end
