#trello-archiver

Simple ruby scripts that you can run to manually or automatically create Excel (XLSX), CSV, and TSV backups of your [Trello](https://trello.com/) boards. //
/
##Install
- Download project
- run
	`bundle install`
- rename `config.example.yml` to `config.yml` and enter your own Trello credentials (see notes in `config.yml` for where & how to get credentials)

##Usage
- Execute one of the scripts in the `bin` directory
- Use the following after authenticating with Ruby-Trello
```ruby
TrelloArchiver::Archiver.new(:format => 'xlsx').create_backup
```

##More info
There are two main scripts you can run that are located in the `bin` directory. *trello_backup.rb* is a command line program that allows you to print a single board at a time from a list of your current boards. *trello_autoarchive.rb* can be used to automatically backup all boards you own at once. You can choose to ignore certain boards by setting their ids into the ignore field in *config.yml*

##Shoutouts
Thanks to Jeremy Tregunna for writing the [ruby-trello library](https://github.com/jeremytregunna/ruby-trello) which made my life a lot easier.

Thanks to mad_typist for initial draft of the code.

##Feedback?
Hit me up at mad_typist[at]yahoo[dot]com
