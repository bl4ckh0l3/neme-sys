<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">

<xsl:template match="neme-sys-tag">
</xsl:template>

<xsl:template match="econeme-sys-tag">
	<xsl:copy-of select="node()"/>
</xsl:template>

<xsl:template match="demoeconeme-sys-tag">
</xsl:template>
