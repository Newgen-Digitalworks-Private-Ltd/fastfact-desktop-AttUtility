Imports System.Security.Cryptography

Public Class clsEncrypt


    Public Class CryptographyFunctions

        Private m_objPublicKeyGenerator As PublicKeyGenerator
        Private m_objRSA As RSACryptoServiceProvider

        Public Sub New()
            MyBase.New()
            m_objPublicKeyGenerator = New PublicKeyGenerator
            m_objRSA = m_objPublicKeyGenerator.GetRSA
            m_objPublicKeyGenerator = Nothing
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
            m_objRSA = Nothing
        End Sub

#Region " Implement Basic Asymmetric (Public/Private Key) Encryption/Decryption Services "

        ' We use Direct Encryption (PKCS#1 v1.5) - so we require MS Windows 2000 or later with high encryption pack installed.

        Public Function SignAndEncrypt(ByRef EncryptionKeyPair As KeyPair, ByVal PlainText As String) As String
            ' Use Public Key to encrypt and private key to sign
            Return Encrypt(EncryptionKeyPair, Sign(EncryptionKeyPair, PlainText))
        End Function

        Public Function Sign(ByRef SigningKeyPair As KeyPair, ByVal Text As String) As String
            'Use PrivateKey to sign
            m_objRSA.FromXmlString(SigningKeyPair.PrivateKey.Key)

            Dim strSignatureData As String

            strSignatureData = ByteArrayAsString(m_objRSA.SignData(StringAsByteArray(Text), _
      System.Security.Cryptography.HashAlgorithm.Create()))

            Return Text & "<signature>" & strSignatureData & "</signature>"
        End Function

        Public Function Encrypt(ByRef EncryptionKeyPair As KeyPair, ByVal PlainText As String) As String
            ' Use Public Key to encrypt
            m_objRSA.FromXmlString(EncryptionKeyPair.PublicKey.Key)

            'Get Modulus Size and compare it to length of PlainText
            ' If Length of PlainText > (Modulus Size - 11), then PlainText will need to be broken into segments of size (Modulus Size - 11)
            ' Each of these segments will be encrypted separately
            '     and will return encrypted strings equal to the Modulus Size (with at least 11 bytes of padding)
            ' When decrypting, if the EncryptedText string > Modulus size, it will be split into segments of size equal to Modulus Size
            ' Each of these EncryptedText segments will be decrypted individually with the resulting PlainText segments re-assembled.

            Dim intBlockSize As Integer = GetModulusSize(EncryptionKeyPair.PublicKey.Key) - 11
            Dim strEncryptedText As String = ""

            While Len(PlainText) > 0
                If Len(PlainText) > intBlockSize Then
                    strEncryptedText = strEncryptedText & EncryptBlock(Left(PlainText, intBlockSize))
                    PlainText = Right(PlainText, Len(PlainText) - intBlockSize)
                Else
                    strEncryptedText = strEncryptedText & EncryptBlock(PlainText)
                    PlainText = ""
                End If
            End While

            Return strEncryptedText
        End Function

        Public Function DecryptAndAuthenticate(ByRef DecryptionKeyPair As KeyPair, ByVal EncryptedText As String) As String
            ' Use Private key to Decrypt and Public Key to Authenticate

            Dim strPlainText As String = ""
            Dim strSignature As String

            strPlainText = Decrypt(DecryptionKeyPair, EncryptedText)

            If Authenticate(DecryptionKeyPair, strPlainText) Then
                strSignature = StripSignature(strPlainText)
                Return strPlainText
            Else
                'Throw new exception
                Throw New Exception("Message authentication failed.")
            End If
        End Function

        Public Function Decrypt(ByRef DecryptionKeyPair As KeyPair, ByVal EncryptedText As String) As String
            ' Use Private Key to Decrypt - don't authenticate

            m_objRSA.FromXmlString(DecryptionKeyPair.PrivateKey.Key)

            ' When decrypting, if the EncryptedText string > Modulus size, it will be split into segments of size equal to Modulus Size
            ' Each of these EncryptedText segments will be decrypted individually with the resulting PlainText segments re-assembled.

            Dim intBlockSize As Integer = GetModulusSize(DecryptionKeyPair.PrivateKey.Key)
            Dim strPlainText As String = ""

            While Len(EncryptedText) > 0
                If Len(EncryptedText) > intBlockSize Then
                    strPlainText = strPlainText & DecryptBlock(Left(EncryptedText, intBlockSize))
                    EncryptedText = Right(EncryptedText, Len(EncryptedText) - intBlockSize)
                Else
                    strPlainText = strPlainText & DecryptBlock(EncryptedText)
                    EncryptedText = ""
                End If
            End While

            Return strPlainText
        End Function

        Public Function Authenticate(ByRef AuthenticationKeyPair As KeyPair, ByVal SignedText As String) As Boolean
            'Use Public Key to Authenticate

            m_objRSA.FromXmlString(AuthenticationKeyPair.PublicKey.Key)

            'Strip Signature from message and use it to validate message

            Dim strSignature As String = StripSignature(SignedText)

            If strSignature <> "" Then
                Return m_objRSA.VerifyData(StringAsByteArray(SignedText), System.Security.Cryptography.HashAlgorithm.Create(), _
          StringAsByteArray(strSignature))
            Else
                Throw New Exception("Digital signature is missing or not formatted properly.")
            End If
        End Function

#End Region

#Region " Helper Functions "

#Region " String <-> Byte Array Conversion Functions "
        Private Function StringAsByteArray(ByVal strIn As String) As Byte()
            Return System.Text.UnicodeEncoding.Default.GetBytes(strIn)
        End Function

        Private Function ByteArrayAsString(ByVal bytesIn As Byte()) As String
            Return System.Text.UnicodeEncoding.Default.GetString(bytesIn)
        End Function
#End Region

#Region " Block level cryptographic functions "
        Private Function EncryptBlock(ByVal strIn As String) As String
            Return ByteArrayAsString(m_objRSA.Encrypt(StringAsByteArray(strIn), False))
        End Function

        Private Function DecryptBlock(ByVal strIn As String) As String
            Return ByteArrayAsString(m_objRSA.Decrypt(StringAsByteArray(strIn), False))
        End Function
#End Region

#Region " Helper cryptographic functions "
        Private Function GetModulusSize(ByVal KeyXml As String) As Integer
            'KeySize is in Bits - so divide by 8 to get # of bytes
            Return m_objRSA.KeySize / 8
        End Function

        Private Function StripSignature(ByRef SignedText As String) As String
            ' Remove SignatureData from SignedText and Return it
            ' Assumption: signature is at end of SignedText and has <signature> and </signature> tags around it

            Dim intStartPosition As Integer
            Dim strSignatureData As String

            intStartPosition = InStr(SignedText, "<signature>", CompareMethod.Text)
            strSignatureData = Right(SignedText, Len(SignedText) - intStartPosition + 1)

            'Strip tags from signature
            strSignatureData = Replace(strSignatureData, "<signature>", "")
            strSignatureData = Replace(strSignatureData, "</signature>", "")

            'Strip signature from SignedText
            SignedText = Replace(SignedText, "<signature>" & strSignatureData & "</signature>", "", , , CompareMethod.Text)

            Return strSignatureData
        End Function
#End Region

#End Region

    End Class ' CryptographyFunctions





    Public Class Key

        Public Enum KeyTypeEnum As Byte
            PublicKey = 0
            PrivateKey = 1
        End Enum

        Private m_strLabel As String
        Private m_strKey As String
        Private m_bytKeyType As KeyTypeEnum

        Public Property Label() As String
            Get
                Label = m_strLabel
            End Get
            Set(ByVal Value As String)
                m_strLabel = Value
            End Set
        End Property

        Public Property Key() As String
            Get
                Key = m_strKey
            End Get
            Set(ByVal Value As String)
                m_strKey = Value
            End Set
        End Property

        Public Property KeyType() As KeyTypeEnum
            Get
                KeyType = m_bytKeyType
            End Get
            Set(ByVal Value As KeyTypeEnum)
                m_bytKeyType = Value
            End Set
        End Property
    End Class ' Key




    Public Class KeyPair

        Private m_strPublicKey As Key
        Private m_strPrivateKey As Key

        Public Property PublicKey() As Key
            Get
                PublicKey = m_strPublicKey
            End Get
            Set(ByVal Value As Key)
                m_strPublicKey = Value
            End Set
        End Property

        Public Property PrivateKey() As Key
            Get
                PrivateKey = m_strPrivateKey
            End Get
            Set(ByVal Value As Key)
                m_strPrivateKey = Value
            End Set
        End Property
    End Class ' KeyPair



    Public Class PublicKeyGenerator

        Public Function GetRSA() As System.Security.Cryptography.RSACryptoServiceProvider
            ' RSA wants to store key info in user's account
            ' we want to use the local (machine) account instead
            Dim objCspParameters As CspParameters = New CspParameters
            objCspParameters.Flags = CspProviderFlags.UseMachineKeyStore
            Dim objRSA As Security.Cryptography.RSACryptoServiceProvider = New Security.Cryptography.RSACryptoServiceProvider(objCspParameters)
            Return objRSA
            objRSA = Nothing
            objCspParameters = Nothing
        End Function

        Public Function MakeKeyPair() As KeyPair
            Dim objRSA As RSA = GetRSA()
            Dim objPublicKey As Key = New Key
            Dim objPrivateKey As Key = New Key
            Dim objKeyPair As New KeyPair

            objKeyPair.PublicKey = objPublicKey
            objKeyPair.PrivateKey = objPrivateKey
            objPublicKey.KeyType = Key.KeyTypeEnum.PublicKey
            objPrivateKey.KeyType = Key.KeyTypeEnum.PrivateKey
            objPublicKey.Key = objRSA.ToXmlString(False)
            objPrivateKey.Key = objRSA.ToXmlString(True)

            Return objKeyPair

            objPublicKey = Nothing
            objPrivateKey = Nothing
            objKeyPair = Nothing
            objRSA = Nothing
        End Function

    End Class ' PublicKeyGenerator





    Public Class KeyRing

        Private m_colKeys As Collection

        Public Sub New()
            MyBase.New()
            m_colKeys = New Collection
        End Sub

        Public Function AddKeyPair(ByRef KeyPair As KeyPair) As KeyPair
            m_colKeys.Add(KeyPair, KeyPair.PublicKey.Label)
            Return m_colKeys.Item(KeyPair.PublicKey.Label)
        End Function

        Public Function GetKeyPair(ByVal PublicKeyLabel As String) As KeyPair
            Return m_colKeys.Item(PublicKeyLabel)
        End Function

        Public Sub RemoveKeyPair(ByVal PublicKeyLabel As String)
            m_colKeys.Remove(PublicKeyLabel)
        End Sub

        Protected Overrides Sub Finalize()
            m_colKeys = Nothing
            MyBase.Finalize()
        End Sub
    End Class ' KeyRing




End Class
