module Sharepoint
  module Publishing
    class VideoChannel < Sharepoint::Object
      include Sharepoint::Type
      sharepoint_resource getter: 'channels', service: 'video'

      def get_video_by_id(id)
        @site.query :get, "#{__metadata['id']}/Videos(guid'#{id}')"
      end
    end

    class VideoItem < Sharepoint::Object
      include Sharepoint::Type
      belongs_to :video_channel

      def get_playback_url(type = 1)
        request "#{__metadata['id']}/GetPlaybackUrl(#{type})"
      end

      def get_token
        request "#{__metadata['id']}/GetStreamingKeyAccessToken"
      end

      private

      def request(url)
        @site.query :get, url, nil, true do |curl|
          curl.headers['Accept'] = 'application/json;odata=nometadata'
          curl.headers['Content-Type'] = curl.headers['Accept']
        end
      end
    end
  end
end
