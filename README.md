# Magic Word Counter

An application written in 2013 that counts words, but with the following restrictions:
- Names, places, and dates count as one word
- Article adjectives are excluded
- MLA citations are excluded

Two versions are included. One is a more feature-rich Windows Forms version. The other is a Windows 8.1 app that was previously published in the Windows store. 
I don't remember if the Windows app has more up to date regular expressions or not.

This is as unmodified from the 2013 version as I could make it, with the following exceptions:
- The debug mode of the Windows 8.1 has been modified to not include trial restrictions.
- The desktop version's reference to DocumentFormat.OpenXml has been updated to use a Nuget package.
