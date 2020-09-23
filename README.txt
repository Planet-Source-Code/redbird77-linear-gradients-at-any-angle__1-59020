[APPINFO]
pGradientDemo.vbp
2005 February 19
redbird77@earthlink.net
http://home.earthlink.net/~redbird77

[ABOUT]
Ho hum, another gradient... but it's pretty fast and supports angles.
I've got some major snippage of a version that supports multiple colors
(each with their own position and transparency).  I'm in the process of
cleaning it all up.  It's pretty fast too since it uses cDIB from Carles PV
and the nifty Bresenham line drawing algorithm.

If you want any snippage, please email me.

[USAGE]
To render a gradient all you need is either:

the DrawGradient sub and the 6 included API declares - OR -
the DrawGradientVB sub and nothing else.

Neither must be in a module, they can go anywhere - form, class, module...

[REVISIONS]
2005 February 19
	Initial release.