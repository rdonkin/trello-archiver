# encoding: utf-8
#!/usr/bin/env ruby

# $LOAD_PATH.unshift 'lib'
require 'trello'
require 'rubygems'
require 'yaml'
require_relative '../lib/trello-archiver.rb'

include Trello
include Trello::Authorization

CONFIG = YAML::load(File.open("config.yml")) unless defined? CONFIG

TrelloArchiver::Authorize.new(CONFIG).authorize

input = TrelloArchiver::Prompt.new(CONFIG).run

TrelloArchiver::Archiver.new(:board => input[:board], :filename => input[:filename], :format => 'tsv').create_backup
