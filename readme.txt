A boolean parser, like that seen in search engines. able to process expressions like:

animals and ((cats or feline) or (dogs or canine)) and "is cuddly"

Written in Perl, Vb, and ASP. That was about as many languages as i could convert it to. Very usefull code.




VB6 INSTALLATION AND USAGE:

simply add the .bas file that comes with this distribution to your visual basic project.

then, loop through a database or whatever, on every record, execute the line:

BlnMatch = checkstring(QUERYSTRING,TEXTTOSEARCH)
If BlnMatch then *display it as a result* else *don't display it*

See the vb6 app that came with this distribution as an example (Project1.vbp)



ASP INSTALLATION AND USAGE:

Simply copy the function 'checkstring' out of the asp file that comes with this distribution, and add it to your search engine, or a library (inlcuded) file for your site.

Usage is exactly the same as VB6, loop through a database or whatever, and:

BlnMatch = checkstring(QUERYSTRING,TEXTTOSEARCH)
If BlnMatch then *display it as a result* else *don't display it*

See the asp page that came with this distribution for an example (search.asp).



PERL INSTALATION AND USAGE:

Copy all the subs to a new perl script (there are 7), and 'require' it from another perl script. Usage is similar to the other two langauges:

$BlnMatch = checkstring($QUERYSTRING,$TEXTTOSEARCH);
if ($BlnMatch eq "true")
{
 *display it as a result*;
}
else
{
 *don't display it*;
}

See the perl page that came with this distribution for an example. Don't forget, on most webservers, you'll need to 'chmod 755' to give the webserver permission to execute it.



All of the examples can be run strait out of the box. They all run on my machine, over my webserver (available on www.planetsourcecode.com)

Scripts are copyright Ashley Harris (ashley___harris@hotmail.com) 2002
You can use them freely in your own apps/scripts.
so long as I get some credit somewhere, and, that I know about it!

    I mean, you can distribute this in any app you want, and make
    obscene amounts of money from it. I just, would like to know
    about it!
    
Also, if you use this, I require a vote on www.planetsourcecode.com, a preaty good deal
if you ask me. (I hide out in the perl, vb, javascript, and asp sections of the site)

Also, if you make an improvement on this script, please let me know what you did!

Ashley

-- 
Ashley Harris
Email : Ashley___harris@hotmail.com
MSN   : Ashley___harris@hotmail.com
ICQ   : 153577070
AIM   : Ashley000Harris
Y!M   : a_s_h_l_e_y_h_a_r_r_i_s