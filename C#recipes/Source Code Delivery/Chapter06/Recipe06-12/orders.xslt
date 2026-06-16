<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    version="1.0" >

  <xsl:template match="Order">
    <html>
      <body>
        <p>
          Order <b>
            <xsl:value-of select="Client/@id"/>
          </b>
          for <xsl:value-of select="Client/Name"/>
        </p>
        <table border="1">
          <td>ID</td>
          <td>Name</td>
          <td>Price</td>
          <xsl:apply-templates select="Items/Item"/>
        </table>
      </body>
    </html>
  </xsl:template>

  <xsl:template match="Items/Item">
    <tr>
      <td>
        <xsl:value-of select="@id"/>
      </td>
      <td>
        <xsl:value-of select="Name"/>
      </td>
      <td>
        <xsl:value-of select="Price"/>
      </td>
    </tr>
  </xsl:template>

</xsl:stylesheet>
