[![Build Status](https://travis-ci.org/boudhayan-dev/OutlookNotify.svg?branch=master)](https://travis-ci.org/boudhayan-dev/OutlookNotify) ![](https://img.shields.io/pypi/pyversions/Django.svg)  ![](https://img.shields.io/pypi/status/Django.svg)


<h1 align="center"> <p align="center"><img src="assets/logo.PNG" /></p> </h1>

<p align="center"> Outlook Notifications made simple ! </p>

<hr/>

<p align="justify">Due to security reasons at my workpalce, it is not allowed to login to Outlook mail from certain devices ( mostly Android ) unless they are screened by the IT department. As a result, there have been times when I would miss important meeting requests and last-minute mails with immediate deadlines because of my absence from my device (work laptop). I had the option of forwarding e-mails to my personal mail to overcome this issue but I decided against it as it would involve the sending of confidential information to an unauthorised email.     

OutlookNotify was developed to solve this problem. It will continuously monitor my system for any new incoming mails and notify me about any last minute meeting requests. It does not forward the entire body of the received email. Instead, it just extracts the timing details of the meeting requests and forwards to my personal email. </p>

<br>

<h3> Download & Installation </h3>

<ul>
  <li>Clone the Github repository.</li>
  <code>git clone https://github.com/boudhayan-dev/OutlookNotify</code>
  <br><br>
  <li>Navigate to the <code>outlook</code> directory</li>
  <code> cd OutlookNotify</code>
  <br><br>
  <li>Install dependencies</li>
  <code> pip install -r requirements.txt</code>
  <br><br>
  <li>Run the script.</li>
  <code>python outlook/main.py</code>
  <br><br>
  <li>Alternatively, run the <code>main.bat</code> file or add it as a task in Windows Task Scheduler to trigger the script when the system in locked.</li>
</ul>

<br>




<h3> Demo</h3>

A sample notification received on my personal mail.
<br>
<table>
    <tr>
      <td>
          <img src="assets/email.PNG">
      </td>
  </tr>
</table>

The following logs demonstrate the mails received during the logging period.
<table>
    <tr>
      <td>
          <img src="assets/logs.PNG">
      </td>
  </tr>
</table>

<br>



<h3>Contributing</h3>
Keep it simple. Keep it minimal. <br>
Also check the issues tab for some enhancement ideas.

<br>


<h3>License</h3>

This project is licensed under the GNU GENERAL PUBLIC LICENSE.
