require 'curb'
require 'json'
require 'sharepoint-session'
require 'sharepoint-object'
require 'sharepoint-types'

module Sharepoint
  class ParseError < StandardError; end
  class RequestsThresholdReached < StandardError; end
  class SPException < StandardError
    def initialize data, uri = nil, body = nil
      @data = data
      @uri  = uri
      @body = body
    end

    def lang         ; @data['error']['message']['lang']  ; end
    def message
      @data['error_description'] || @data['error']['message']['value']
    end
    def code         ; @data['error']['code'] ; end
    def uri          ; @uri ; end
    def request_body ; @body ; end
  end

  class Site
    attr_reader   :server_url
    attr_accessor :url, :protocol
    attr_accessor :session
    attr_accessor :name
    attr_accessor :verbose

    def initialize server_url, site_name
      @server_url  = server_url
      @name        = site_name
      @url         = "#{@server_url}/#{@name}"
      @session     = Session.new self
      @web_context = nil
      @protocol    = 'https'
      @verbose     = false
    end

    def authentication_path
      "#{@protocol}://#{@server_url}/_forms/default.aspx?wa=wsignin1.0"
    end

    def api_path uri, service
      if service == 'video'
        "#{@protocol}://#{@server_url}/portals/hub/_api/videoservice/#{uri}"
      else
        "#{@protocol}://#{@url}/_api/web/#{uri}"
      end
    end

    def filter_path uri
      uri
    end

    def context_info
      query :get, ''
    end

    # Sharepoint uses 'X-RequestDigest' as a CSRF security-like.
    # The form_digest method acquires a token or uses a previously acquired
    # token if it is still supposed to be valid.
    def form_digest
      if @web_context.nil? or (not @web_context.is_up_to_date?)
        @getting_form_digest = true
        @web_context         = query :post, "#{@protocol}://#{@url}/_api/contextinfo"
        @getting_form_digest = false
      end
      @web_context.form_digest_value
    end

    def query method, uri, body = nil, skip_json=false, service = 'web', &block
      uri        = uri =~ /^http/ ? uri : api_path(uri, service)
      arguments  = [ uri ]
      arguments << body unless method == :get
      Rails.logger.info("Sharepoint:query before send #{method}->#{uri}(#{body})")
      result = Curl::Easy.send "http_#{method}", *arguments do |curl|
        curl.headers["Cookie"]          = @session.cookie
        curl.headers["Accept"]          = "application/json;odata=verbose"
        if method != :get
          curl.headers["Content-Type"]    = curl.headers["Accept"]
          unless @getting_form_digest
            curl.headers["X-RequestDigest"] = form_digest
            Rails.logger.info("Sharepoint:query get digest #{method} #{uri} RequestDigest #{form_digest}")
          end
        end
        curl.verbose = @verbose
        curl.ssl_verify_peer = false
        @session.send :curl, curl if @session.methods.include? :curl
        block.call curl           if block.present?
      end

      if skip_json || result.body_str.nil? || result.body_str.empty?
        result.body_str
      else
        begin
          data = JSON.parse result.body_str
          Rails.logger.info("Sharepoint:query #{uri} returned #{data}")
          error = data['error'] || data['error_description']
          if error
            Rails.logger.info("Sharepoint:query receive error #{error.inspect} from #{method}->#{uri}(#{body})")
            raise Sharepoint::SPException.new data, uri, body
          end

          make_object_from_response data
        rescue JSON::ParserError => e
          Rails.logger.info("Sharepoint:query Sharepoint::ParseError from #{method}->#{uri}(#{body}) resulted in body=#{body}, e=#{e.inspect}, #{e.backtrace.inspect}, response=#{result.body_str}")
          raise Sharepoint::RequestsThresholdReached if result.response_code == 429
          raise Sharepoint::ParseError.new("Sharepoint::ParseError with body=#{body}, e=#{e.inspect}, #{e.backtrace.inspect}, response=#{result.body_str}")
        end
      end
    end

    def make_object_from_response data
      if data['d']['results'].nil?
        data['d'] = data['d'][data['d'].keys.first] if data['d']['__metadata'].nil?
        if not data['d'].nil?
          make_object_from_data data['d']
        else
          nil
        end
      else
        array = Array.new
        data['d']['results'].each do |result|
          array << (make_object_from_data result)
        end
        array
      end
    end

    # Uses sharepoint's __metadata field to solve which Ruby class to instantiate,
    # and return the corresponding Sharepoint::Object.
    def make_object_from_data data
      type_name  = data['__metadata']['type'].gsub(/(^(SP\.|(\w+\.)+)|\(.*)/, '')
                                             .gsub(/^Collection\(Edm\.String\)/, 'CollectionString')
                                             .gsub(/^Collection\(Edm\.Int32\)/, 'CollectionInteger')
      type_parts = type_name.split '.'
      type_name  = type_parts.pop
      constant   = Sharepoint
      type_parts.each do |part| constant = constant.const_get part end

      klass      = constant.const_get type_name rescue nil
      if klass
        klass.new self, data
      else
        Sharepoint::GenericSharepointObject.new type_name, self, data
      end
    end
  end
end
