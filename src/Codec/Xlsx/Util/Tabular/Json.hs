{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}
-- | Convinience utility to decode Xlsx tabular.
module Codec.Xlsx.Util.Tabular.Json
       ( parseJSON
       , toJSON
       ) where

import Codec.Xlsx.Util.Tabular.Imports
import Data.Char (toLower)


instance ToJSON RichTextRun where
  toJSON r =
    object [ "text" .= view richTextRunText r ]

instance FromJSON RichTextRun where
  parseJSON (Object v) =
    RichTextRun <$> pure Nothing <*> (v .: "text")


deriveJSON defaultOptions
  { fieldLabelModifier = drop 5
  , constructorTagModifier = map toLower . drop 4
  } ''CellValue


deriveJSON defaultOptions
  { fieldLabelModifier = map toLower . drop 8
  , constructorTagModifier = map toLower
  } ''TabularRow


deriveJSON defaultOptions
  { fieldLabelModifier = map toLower . drop 12
  , constructorTagModifier = map toLower
  } ''TabularHead


deriveJSON defaultOptions
  { fieldLabelModifier = map toLower . drop 8
  , constructorTagModifier = map toLower
  } ''Tabular
