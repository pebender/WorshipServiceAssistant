function OpenNewWindow(url)
{
   newWindow = window.open(url);
   if (navigator.appName == 'Netscape')
   {
      newWindow.focus();
   }
}
