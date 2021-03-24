
const Imap = require('imap');
const fs = require('fs');
const jsdom = require("jsdom");
const simpleParser = require('mailparser').simpleParser;

const filename = 'idmsb.xlsx';
const mails = [];

const imap = new Imap({
  user: 'EMAIL_NAME',
  password: 'EMAIL_PASS (password from aplications pass)',
  host: 'imap.yandex.ru',
  port: 993,
  tls: true
});

function openInbox(cb) {
  imap.openBox('INBOX', true, cb);
}

imap.once('ready', function() {
  openInbox(function(err, box) {
    if (err) throw err;
    var f = imap.seq.fetch('1:1', {
      bodies: ''
    });
    
    // search example
   imap.search([ 'ALL', ['SUBJECT', 'IDMSB (Ипотека)'] ], function(err, results) {
     if (err) throw err;
     var f = imap.fetch(results, { bodies: '' });

        f.on('message', function (msg, seqno) {
            msg.on('body', function (stream, info) {
                simpleParser(stream, (err, parsed) => {
                    if (parsed.html) {
                        const dom = new jsdom.JSDOM(parsed.html);
                        
                        const person = {
                            name: dom.window.document.querySelector('p:nth-child(3) > b').textContent,
                            phone: dom.window.document.querySelector('p:nth-child(4) > b').textContent,
                        }
                        const link = dom.window.document.querySelector('p:nth-child(5)').textContent.split('-')[1];
                        
                        if (link) {
                            const queryString = link.split('?')[1];
                            const queryArr = queryString ? queryString.split('&') : [];
                            queryArr.forEach(item => {
                                const arr = item.split('=');
                                
                                person[arr[0]] = arr[1];
                            });
                        }

                        mails.push(person);

                        console.log(person);   
                    }
                });
            });
        });

        f.once('end', function() {
            fs.writeFileSync('idmsb.json', JSON.stringify(mails));

            console.log('done!');

            imap.end();
        });
     
   });

  });
});

imap.connect();